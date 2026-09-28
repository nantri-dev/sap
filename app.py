"""
LedgerMind - local web briefing (no Power BI needed).

Run from the project folder (same folder as main.py):
    python app.py
Then open http://127.0.0.1:5000  (it also opens automatically).

What it does
  * Reads output/briefing.csv (columns: category, title, detail, priority)
  * Shows it as clickable cards, sorted so HIGH priority is always on top
  * Click a card  -> the local model (Ollama) explains it and offers follow-up questions
  * Kill switch   -> turns the AI off; the page falls back to raw rule-flagged data
  * Memory badges -> each item is tagged NEW or SEEN BEFORE (with a run count)
  * Upload button -> load ANY Excel/CSV; universal_ingest.py works out the columns itself
  * Refresh button-> re-runs the pipeline (your uploaded file, or main.py for the sample data)
  * Ask bar       -> type any question. Keyword search always works (even with the AI off);
                     when the AI is ON the local model also writes an answer

Other devices (phone, second laptop on the same Wi-Fi):
    Windows PowerShell:  $env:LEDGERMIND_HOST="0.0.0.0"; python app.py
    Mac / Linux:         LEDGERMIND_HOST=0.0.0.0 python3 app.py
  A private access token is created automatically and printed as a link.
Everything stays on this machine.
"""
import csv
import json
import os
import re
import secrets
import shutil
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from datetime import datetime

import requests
from flask import Flask, Response, g, jsonify, request
from werkzeug.utils import secure_filename

import universal_ingest

BASE = os.path.dirname(os.path.abspath(__file__))
WORKSPACES = os.path.join(BASE, "workspaces")
SAMPLE_XLSX = os.path.join(BASE, "sample", "sample_data.xlsx")   # used to seed public-demo visitors
PUBLIC = os.environ.get("LEDGERMIND_PUBLIC", "") == "1"          # 1 = hosted demo, one workspace per visitor


def ws_root():
    """Folder holding this visitor's data. On your laptop it is just the project folder."""
    return getattr(g, "ws_root", BASE)


def p_briefing():
    return os.path.join(ws_root(), "output", "briefing.csv")


def p_summary():
    return os.path.join(ws_root(), "output", "data_summary.txt")


def p_data():
    return os.path.join(ws_root(), "data")


def p_seen():
    return os.path.join(ws_root(), "memory", "web_seen.json")      # powers the memory badges


def p_switch():
    return os.path.join(ws_root(), "memory", "ai_switch.json")     # powers the kill switch


def p_active():
    return os.path.join(ws_root(), "memory", "active_file.json")   # the spreadsheet in use


OLLAMA_BASE = os.environ.get("OLLAMA_BASE", "http://localhost:11434")
OLLAMA_URL = OLLAMA_BASE + "/api/generate"
MODEL = os.environ.get("LEDGERMIND_MODEL", "llama3.2:1b")
PORT = int(os.environ.get("PORT", "5000"))
HOST = os.environ.get("LEDGERMIND_HOST", "127.0.0.1")
TOKEN = os.environ.get("LEDGERMIND_TOKEN", "")

PRIORITY_RANK = {"high": 0, "medium": 1, "low": 2}

GUARDRAILS = (
    "You are LedgerMind, a local finance briefing assistant. "
    "Use ONLY the data given below. Never invent numbers, dates, counterparties or accounts. "
    "If something is not in the data, say it is not in the data. "
    "You only explain and suggest what a human should review; "
    "you never approve, execute or move money."
)

STOPWORDS = {"the", "and", "for", "what", "which", "how", "are", "was", "were", "this", "that", "with",
             "any", "all", "has", "have", "from", "about", "show", "tell", "why", "who", "when",
             "there", "our", "does", "did"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # uploads up to 25 MB


# --------------------------------------------------------------------------
# access token (only enforced for other devices; this laptop is always allowed)
# --------------------------------------------------------------------------
@app.before_request
def require_token():
    if not TOKEN or request.remote_addr in ("127.0.0.1", "::1"):
        return None
    if request.args.get("token") == TOKEN or request.cookies.get("lm_token") == TOKEN:
        return None
    return Response("Access token required. Open the link printed in the LedgerMind window "
                    "(it ends with ?token=...).", status=401)


@app.after_request
def remember_token(resp):
    if TOKEN and request.args.get("token") == TOKEN:
        resp.set_cookie("lm_token", TOKEN, httponly=True, samesite="Lax")
    if getattr(g, "new_sid", None):  # SameSite=None so it also works when the page is shown inside an iframe
        resp.set_cookie("lm_sid", g.new_sid, httponly=True, samesite="None", secure=True, max_age=86400)
    return resp


def _cleanup_workspaces(max_age_hours=24):
    if not os.path.isdir(WORKSPACES):
        return
    now = time.time()
    for name in os.listdir(WORKSPACES):
        path = os.path.join(WORKSPACES, name)
        if os.path.isdir(path) and now - os.path.getmtime(path) > max_age_hours * 3600:
            shutil.rmtree(path, ignore_errors=True)


def _seed_workspace(root):
    """First visit in public mode: give the visitor their own copy of the sample briefing."""
    _cleanup_workspaces()
    os.makedirs(os.path.join(root, "data"), exist_ok=True)
    os.makedirs(os.path.join(root, "memory"), exist_ok=True)
    if os.path.exists(SAMPLE_XLSX):
        target = os.path.join(root, "data", "sample_data.xlsx")
        shutil.copyfile(SAMPLE_XLSX, target)
        try:
            universal_ingest.run(target, root=root)
            _save_json(os.path.join(root, "memory", "active_file.json"),
                       {"path": target, "name": "sample_data.xlsx (made-up demo data)",
                        "uploaded_at": datetime.now().isoformat(timespec="seconds")})
        except Exception as exc:  # noqa: BLE001
            print("Could not seed sample data:", exc)


@app.before_request
def pick_workspace():
    if not PUBLIC:
        g.ws_root = BASE
        return None
    sid = request.cookies.get("lm_sid", "")
    if not re.fullmatch(r"[0-9a-f]{16}", sid):
        sid = secrets.token_hex(8)
        g.new_sid = sid
    g.ws_root = os.path.join(WORKSPACES, sid)
    if not os.path.isdir(g.ws_root):
        _seed_workspace(g.ws_root)
    return None


# --------------------------------------------------------------------------
# small helpers
# --------------------------------------------------------------------------
def _load_json(path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return default


def _save_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


# --------------------------------------------------------------------------
# kill switch
# --------------------------------------------------------------------------
def ai_enabled():
    return bool(_load_json(p_switch(), {"ai_enabled": True}).get("ai_enabled", True))


def set_ai_enabled(value):
    _save_json(p_switch(), {
        "ai_enabled": bool(value),
        "changed_at": datetime.now().isoformat(timespec="seconds"),
    })


# --------------------------------------------------------------------------
# briefing data
# --------------------------------------------------------------------------
def load_items():
    """Read briefing.csv, tidy it, drop exact duplicate rows, put HIGH first."""
    if not os.path.exists(p_briefing()):
        return []
    items, seen = [], set()
    with open(p_briefing(), "r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            category = (row.get("category") or "").strip().lower()
            title = (row.get("title") or "").strip()
            detail = (row.get("detail") or "").strip()
            priority = (row.get("priority") or "low").strip().lower()
            if not title:
                continue
            if priority not in PRIORITY_RANK:
                priority = "low"
            detail = re.sub(r"due\s+nan", "no due date on file", detail, flags=re.I)
            key = f"{category}|{title}|{detail}".lower()
            if key in seen:
                continue
            seen.add(key)
            items.append({"key": key, "category": category, "title": title,
                          "detail": detail, "priority": priority})
    items.sort(key=lambda i: PRIORITY_RANK[i["priority"]])  # stable sort
    for idx, item in enumerate(items):
        item["idx"] = idx
    return items


def data_summary():
    """Compact facts about the WHOLE uploaded file (written by universal_ingest.py)."""
    try:
        with open(p_summary(), "r", encoding="utf-8") as f:
            return f.read()[:4500]
    except OSError:
        return ""


def update_history(items):
    """
    The 'memory' badges. Every time briefing.csv changes (a new pipeline run)
    we count how many runs each item has appeared in.
    Returns (history_by_key, titles_that_disappeared_since_last_run).
    """
    try:
        run_id = os.path.getmtime(p_briefing())
    except OSError:
        return {}, []
    state = _load_json(p_seen(), {"last_run": None, "items": {}, "last_keys": [], "dropped": []})
    if state.get("last_run") != run_id:
        now = datetime.now().isoformat(timespec="seconds")
        current_keys = [i["key"] for i in items]
        dropped = [state["items"][k]["title"] for k in state.get("last_keys", [])
                   if k not in current_keys and k in state["items"]]
        for item in items:
            rec = state["items"].get(item["key"])
            if rec:
                rec["times_seen"] += 1
                rec["last_seen"] = now
            else:
                state["items"][item["key"]] = {"title": item["title"], "first_seen": now,
                                               "last_seen": now, "times_seen": 1}
        state["last_run"] = run_id
        state["last_keys"] = current_keys
        state["dropped"] = dropped
        _save_json(p_seen(), state)
    return state["items"], state.get("dropped", [])


# --------------------------------------------------------------------------
# local AI (Ollama)
# --------------------------------------------------------------------------
def ask_ollama(prompt, want_json=False, timeout=120):
    payload = {"model": MODEL, "prompt": prompt, "stream": False,
               "options": {"temperature": 0.2}}
    if want_json:
        payload["format"] = "json"
    try:
        r = requests.post(OLLAMA_URL, json=payload, timeout=timeout)
        r.raise_for_status()
        return r.json().get("response", "").strip(), None
    except requests.exceptions.ConnectionError:
        return None, "Could not reach Ollama. Is it running? Open http://localhost:11434 to check."
    except requests.exceptions.Timeout:
        return None, "The local model took too long to answer. Try again in a moment."
    except Exception as exc:  # noqa: BLE001
        return None, f"Local model error: {exc}"


def briefing_context(items):
    return "\n".join(
        f"- [{i['priority'].upper()}] ({i['category']}) {i['title']}: {i['detail']}" for i in items
    )


def item_line(item):
    return f"[{item['priority'].upper()}] ({item['category']}) {item['title']}: {item['detail']}"


def plain_text(item):
    return (f"{item['title']}: {item['detail']} Priority: {item['priority']}. "
            f"Category: {item['category'] or 'n/a'}. "
            "This is the rule-flagged data exactly as the pipeline produced it.")


def find_item(items, idx):
    try:
        return items[int(idx)]
    except (ValueError, TypeError, IndexError):
        return None


def keyword_search(items, question):
    words = [w for w in re.findall(r"[a-z0-9]+", question.lower())
             if len(w) > 2 and w not in STOPWORDS]
    scored = []
    for i in items:
        hay = f"{i['category']} {i['title']} {i['detail']} {i['priority']}".lower()
        score = sum(1 for w in words if w in hay)
        if score:
            scored.append((score, i))
    scored.sort(key=lambda s: (-s[0], PRIORITY_RANK[s[1]["priority"]]))
    return [i for _, i in scored[:5]]


def summary_block():
    s = data_summary()
    return ("\n\nFacts about the whole data file:\n" + s) if s else ""


# --------------------------------------------------------------------------
# API
# --------------------------------------------------------------------------
@app.get("/api/briefing")
def api_briefing():
    items = load_items()
    history, dropped = update_history(items)
    out = []
    for item in items:
        rec = history.get(item["key"], {})
        out.append({
            "idx": item["idx"], "category": item["category"], "title": item["title"],
            "detail": item["detail"], "priority": item["priority"],
            "times_seen": rec.get("times_seen", 1), "first_seen": rec.get("first_seen", ""),
        })
    counts = {p: sum(1 for i in items if i["priority"] == p) for p in PRIORITY_RANK}
    counts["total"] = len(items)
    generated = ""
    if os.path.exists(p_briefing()):
        generated = datetime.fromtimestamp(os.path.getmtime(p_briefing())).strftime("%d %b %Y, %I:%M %p")
    active = _load_json(p_active(), {})
    return jsonify({"items": out, "counts": counts, "dropped": dropped,
                    "generated_at": generated, "ai_enabled": ai_enabled(),
                    "source": active.get("name", ""), "public": PUBLIC})


@app.post("/api/switch")
def api_switch():
    data = request.get_json(silent=True) or {}
    value = data.get("ai_enabled")
    set_ai_enabled((not ai_enabled()) if value is None else bool(value))
    return jsonify({"ai_enabled": ai_enabled()})


@app.post("/api/expand")
def api_expand():
    items = load_items()
    data = request.get_json(silent=True) or {}
    item = find_item(items, data.get("idx"))
    if item is None:
        return jsonify({"error": "Item not found. Reload the page."}), 404

    if not ai_enabled():
        return jsonify({"ai": False, "explanation": plain_text(item), "followups": []})

    prompt = (
        GUARDRAILS
        + "\n\nToday's full briefing:\n" + briefing_context(items)
        + "\n\nThe user opened this item: " + item_line(item)
        + '\n\nReply as JSON with exactly two keys: "explanation" (2-3 plain-English sentences '
          'on what this item means and why it may matter) and "followups" (a list of 3 short '
          "follow-up questions the user might ask about THIS item, answerable from the data)."
    )
    text, err = ask_ollama(prompt, want_json=True)
    if err:
        return jsonify({"ai": True, "error": err, "explanation": plain_text(item), "followups": []})

    explanation, followups = text, []
    try:
        parsed = json.loads(text)
        explanation = str(parsed.get("explanation", "")).strip() or text
        raw = parsed.get("followups", [])
        if isinstance(raw, list):
            followups = [str(q).strip() for q in raw if str(q).strip()][:3]
    except (json.JSONDecodeError, AttributeError):
        pass
    if not followups:
        followups = ["Why was this flagged?",
                     "How does this compare with the other items today?",
                     "What should a human check next?"]
    return jsonify({"ai": True, "explanation": explanation, "followups": followups})


@app.post("/api/answer")
def api_answer():
    items = load_items()
    data = request.get_json(silent=True) or {}
    item = find_item(items, data.get("idx"))
    question = (data.get("question") or "").strip()[:300]
    if item is None or not question:
        return jsonify({"error": "Missing item or question."}), 400
    if not ai_enabled():
        return jsonify({"ai": False, "answer": "AI reasoning is switched off (kill switch), "
                                               "so no answer is generated. Raw data is shown above."})
    prompt = (
        GUARDRAILS
        + "\n\nToday's full briefing:\n" + briefing_context(items)
        + summary_block()
        + "\n\nItem in focus: " + item_line(item)
        + "\n\nQuestion: " + question
        + "\n\nAnswer in at most 3 short sentences."
    )
    text, err = ask_ollama(prompt)
    if err:
        return jsonify({"ai": True, "error": err})
    return jsonify({"ai": True, "answer": text})


@app.post("/api/ask")
def api_ask():
    items = load_items()
    data = request.get_json(silent=True) or {}
    question = (data.get("question") or "").strip()[:300]
    if not question:
        return jsonify({"error": "Type a question first."}), 400
    matches = [{"idx": i["idx"], "priority": i["priority"], "title": i["title"],
                "detail": i["detail"]} for i in keyword_search(items, question)]
    if not ai_enabled():
        return jsonify({"ai": False, "answer": "", "matches": matches})
    prompt = (
        GUARDRAILS
        + "\n\nToday's full briefing:\n" + briefing_context(items)
        + summary_block()
        + "\n\nQuestion: " + question
        + "\n\nAnswer in at most 4 short sentences. If the data does not contain "
          "the answer, say so plainly."
    )
    text, err = ask_ollama(prompt)
    return jsonify({"ai": True, "answer": text or "", "error": err, "matches": matches})


@app.post("/api/upload")
def api_upload():
    """Accept ANY .xlsx/.xlsm/.csv, work out its columns, rebuild the briefing."""
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"ok": False, "error": "No file received."}), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".csv"):
        return jsonify({"ok": False, "error": "Please upload an .xlsx, .xlsm or .csv file "
                                              "(old .xls: open it in Excel and Save As .xlsx)."}), 400
    os.makedirs(p_data(), exist_ok=True)
    name = secure_filename(f.filename) or ("uploaded" + ext)
    path = os.path.join(p_data(), name)
    f.save(path)
    try:
        report = universal_ingest.run(path, root=ws_root())
    except Exception as exc:  # noqa: BLE001
        return jsonify({"ok": False, "error": f"Could not read that file: {exc}"}), 400
    _save_json(p_active(), {"path": path, "name": name,
                             "uploaded_at": datetime.now().isoformat(timespec="seconds")})
    return jsonify({"ok": True, "report": report})


@app.post("/api/refresh")
def api_refresh():
    """Re-runs the pipeline so the demo can be done from the browser."""
    active = _load_json(p_active(), {})
    if active.get("path") and os.path.exists(active["path"]):
        try:
            universal_ingest.run(active["path"], root=ws_root())
        except Exception as exc:  # noqa: BLE001
            return jsonify({"ok": False, "error": f"Could not read {active.get('name')}: {exc}"}), 500
        return jsonify({"ok": True})
    if PUBLIC:
        return jsonify({"ok": False, "error": "Nothing to refresh yet. Upload a file first."}), 400
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    try:
        result = subprocess.run([sys.executable, os.path.join(BASE, "main.py")], cwd=BASE,
                                capture_output=True, encoding="utf-8", errors="replace",
                                timeout=600, env=env)
    except subprocess.TimeoutExpired:
        return jsonify({"ok": False, "error": "The pipeline took longer than 10 minutes."}), 500
    if result.returncode != 0:
        tail = (result.stderr or result.stdout or "").strip().splitlines()[-5:]
        return jsonify({"ok": False, "error": "main.py failed: " + " | ".join(tail)}), 500
    return jsonify({"ok": True})


# --------------------------------------------------------------------------
# page
# --------------------------------------------------------------------------
PAGE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>LedgerMind - Morning Briefing</title>
<style>
  :root{--bg:#0B0A20;--panel:#15132F;--panel2:#1D1A42;--text:#F4F7FF;--muted:#A8B0C8;
        --purple:#8B5CF6;--cyan:#22D3EE;--gold:#FBBF24;--red:#F87171;--green:#34D399;}
  *{box-sizing:border-box}
  body{margin:0;min-height:100vh;color:var(--text);
       background:radial-gradient(1200px 600px at 80% -10%,#231a5a 0%,var(--bg) 60%);
       font-family:"Segoe UI",system-ui,-apple-system,Arial,sans-serif}
  .wrap{max-width:920px;margin:0 auto;padding:28px 20px 60px}
  header{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;flex-wrap:wrap}
  h1{margin:0;font-size:28px}
  h1 span{color:var(--gold)}
  .sub{color:var(--muted);margin-top:6px;font-size:14px}
  .controls{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
  button{font:inherit;cursor:pointer;border-radius:999px;border:1px solid #3a3670;
         background:var(--panel2);color:var(--text);padding:8px 14px}
  button:hover{border-color:var(--purple)}
  button:disabled{opacity:.6;cursor:wait}
  .sw.on{border-color:var(--green);color:var(--green)}
  .sw.off{border-color:var(--red);color:var(--red);background:#2a1620}
  .banner{margin:18px 0 0;padding:12px 14px;border-radius:12px;background:#2a1620;
          border:1px solid var(--red);color:#ffd5d5;font-size:14px}
  .stats{display:flex;gap:10px;flex-wrap:wrap;margin:22px 0 8px}
  .chip{background:var(--panel);border:1px solid #2b2860;border-radius:12px;padding:10px 14px;min-width:96px}
  .chip b{display:block;font-size:22px}
  .chip small{color:var(--muted)}
  .chip.high b{color:var(--red)} .chip.medium b{color:var(--gold)} .chip.low b{color:var(--cyan)}
  h2{font-size:13px;letter-spacing:.14em;text-transform:uppercase;color:var(--muted);margin:26px 0 10px}
  .card{background:var(--panel);border:1px solid #2b2860;border-left:5px solid var(--cyan);
        border-radius:14px;margin-bottom:12px;overflow:hidden}
  .card.p-high{border-left-color:var(--red)} .card.p-medium{border-left-color:var(--gold)}
  .card-head{padding:14px 16px;cursor:pointer}
  .card-head:hover{background:var(--panel2)}
  .badges{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px}
  .badge{font-size:11px;letter-spacing:.08em;text-transform:uppercase;padding:3px 8px;
         border-radius:999px;border:1px solid #3a3670;color:var(--muted)}
  .badge.high{color:var(--red);border-color:var(--red)}
  .badge.medium{color:var(--gold);border-color:var(--gold)}
  .badge.low{color:var(--cyan);border-color:var(--cyan)}
  .badge.new{color:var(--gold);border-color:var(--gold);background:#2a2410}
  .badge.seen{color:#c4b5fd;border-color:var(--purple);background:#211a4a}
  .title{font-weight:600;font-size:17px}
  .detail{color:var(--muted);margin-top:4px;font-size:14px}
  .hint{color:#6f77a0;font-size:12px;margin-top:8px}
  .card-body{padding:4px 16px 16px;border-top:1px solid #2b2860;background:#110f2b}
  .card-body[hidden],.answer[hidden]{display:none}
  .label{font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:var(--purple);margin:12px 0 6px}
  .text{line-height:1.5;font-size:14.5px}
  .err{color:#ffb4b4;font-size:13px;margin-top:8px}
  .fups{display:flex;flex-direction:column;gap:8px;margin-top:8px}
  .fup{text-align:left;border-radius:10px;font-size:14px}
  .answer{margin:6px 0 4px 10px;padding-left:10px;border-left:2px solid var(--purple);
          font-size:14px;line-height:1.5;color:#dfe4ff}
  .dropped{background:var(--panel);border:1px dashed #3a3670;border-radius:12px;
           padding:12px 14px;color:var(--muted);font-size:14px}
  footer{margin-top:34px;color:#6f77a0;font-size:12px;text-align:center}
  .empty{padding:30px;text-align:center;color:var(--muted)}
  .ask{display:flex;gap:8px;margin:22px 0 6px}
  .ask input{flex:1;min-width:0;padding:12px 16px;border-radius:999px;border:1px solid #3a3670;background:var(--panel);color:var(--text);font:inherit}
  .askbox{background:var(--panel);border:1px solid var(--purple);border-radius:12px;padding:14px;margin-top:8px;line-height:1.5;font-size:14.5px}
  .note{margin:14px 0 0;padding:10px 14px;border-radius:12px;background:#2a2410;border:1px solid var(--gold);color:#ffe9a8;font-size:13.5px}
  .match{padding:6px 0;border-top:1px solid #2b2860;font-size:14px;color:var(--muted)}
  @media (max-width:560px){h1{font-size:22px}.wrap{padding:18px 14px 50px}}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div>
      <h1>Ledger<span>Mind</span> - Morning Briefing</h1>
      <div class="sub" id="sub">Loading...</div>
    </div>
    <div class="controls">
      <input type="file" id="file" accept=".xlsx,.xlsm,.csv" hidden>
      <button id="upload">&#8593; Upload your data</button>
      <button id="refresh">&#8635; Refresh briefing</button>
      <button id="switch" class="sw on">AI reasoning: ON</button>
    </div>
  </header>
  <div id="banner"></div>
  <div id="publicnote"></div>
  <div id="uploadout"></div>
  <div class="ask">
    <input id="q" placeholder="Ask anything about today's briefing...">
    <button id="askbtn">Ask</button>
  </div>
  <div id="askout"></div>
  <div class="stats" id="stats"></div>
  <h2>Needs your attention</h2>
  <div id="list"></div>
  <div id="droppedWrap"></div>
  <footer id="foot">Runs entirely on this machine &middot; local model via Ollama &middot; no data leaves this device &middot; the AI explains, you decide</footer>
</div>
<script>
const $ = (s) => document.querySelector(s);
const el = (tag, cls, text) => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
};
const post = (url, body) => fetch(url, {
  method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(body || {})
}).then(r => r.json());

let state = {items: [], counts: {}, dropped: [], ai_enabled: true, generated_at: "", source: ""};

function greeting() {
  const h = new Date().getHours();
  return h < 12 ? "Good morning" : h < 17 ? "Good afternoon" : "Good evening";
}

function paintSwitch() {
  const b = $("#switch");
  b.className = "sw " + (state.ai_enabled ? "on" : "off");
  b.textContent = state.ai_enabled ? "AI reasoning: ON" : "AI reasoning: OFF (kill switch)";
  const banner = $("#banner");
  banner.innerHTML = "";
  if (!state.ai_enabled) {
    banner.appendChild(el("div", "banner",
      "Kill switch is active. The AI is disabled: this page shows rule-flagged data only, with no AI-written text."));
  }
}

function card(it) {
  const c = el("div", "card p-" + it.priority);
  const head = el("div", "card-head");
  const badges = el("div", "badges");
  badges.appendChild(el("span", "badge " + it.priority, it.priority));
  if (it.category) badges.appendChild(el("span", "badge", it.category));
  if (it.times_seen > 1) {
    const s = el("span", "badge seen", "seen before \u00b7 " + it.times_seen + " runs");
    s.title = "First seen " + it.first_seen;
    badges.appendChild(s);
  } else {
    badges.appendChild(el("span", "badge new", "new"));
  }
  head.appendChild(badges);
  head.appendChild(el("div", "title", it.title));
  head.appendChild(el("div", "detail", it.detail));
  head.appendChild(el("div", "hint", "Click to expand \u25be"));
  const body = el("div", "card-body");
  body.hidden = true;
  let loaded = false;
  head.onclick = async () => {
    body.hidden = !body.hidden;
    if (!body.hidden && !loaded) { loaded = true; await fill(body, it); }
  };
  c.appendChild(head);
  c.appendChild(body);
  return c;
}

async function fill(body, it) {
  body.appendChild(el("div", "label", state.ai_enabled ? "Thinking locally..." : "Raw data"));
  let d;
  try { d = await post("/api/expand", {idx: it.idx}); }
  catch (e) { body.innerHTML = ""; body.appendChild(el("div", "err", "Could not reach the local server.")); return; }
  body.innerHTML = "";
  if (d.error && !d.explanation) { body.appendChild(el("div", "err", d.error)); return; }
  body.appendChild(el("div", "label", (d.ai && !d.error) ? "AI explanation (local model)" : "Rule-based summary (AI off)"));
  body.appendChild(el("div", "text", d.explanation));
  if (d.error) body.appendChild(el("div", "err", d.error));
  if (d.followups && d.followups.length) {
    body.appendChild(el("div", "label", "Dig deeper"));
    const box = el("div", "fups");
    d.followups.forEach(q => {
      const wrap = el("div");
      const b = el("button", "fup", q);
      const ans = el("div", "answer");
      ans.hidden = true;
      b.onclick = async () => {
        ans.hidden = false;
        ans.textContent = "Thinking locally...";
        try {
          const r = await post("/api/answer", {idx: it.idx, question: q});
          ans.textContent = r.answer || r.error || "No answer.";
        } catch (e) { ans.textContent = "Could not reach the local server."; }
      };
      wrap.appendChild(b);
      wrap.appendChild(ans);
      box.appendChild(wrap);
    });
    body.appendChild(box);
  }
}

function render() {
  paintSwitch();
  if (state.public) $("#foot").textContent = "Demo instance with an open-source model running inside this same container \u00b7 no AI vendor is called \u00b7 in production the same container runs inside your own network \u00b7 the AI explains, you decide";
  const pn = $("#publicnote");
  pn.innerHTML = "";
  if (state.public) pn.appendChild(el("div", "note",
    "Public demo: this workspace is private to your browser and is deleted after 24 hours. Please use made-up data only."));
  const c = state.counts || {};
  const src = state.source ? " Source file: " + state.source + "." : "";
  $("#sub").textContent = state.generated_at
    ? greeting() + ". Briefing generated " + state.generated_at + "." + src
    : greeting() + ". No briefing yet - click Upload your data.";
  const stats = $("#stats");
  stats.innerHTML = "";
  [["total", "Items"], ["high", "High"], ["medium", "Medium"], ["low", "Low"]].forEach(([k, label]) => {
    const chip = el("div", "chip " + k);
    chip.appendChild(el("b", null, String(c[k] || 0)));
    chip.appendChild(el("small", null, label));
    stats.appendChild(chip);
  });
  const list = $("#list");
  list.innerHTML = "";
  if (!state.items.length) {
    list.appendChild(el("div", "empty", "Nothing to show yet. Click Upload your data (Excel or CSV) or Refresh briefing."));
  }
  state.items.forEach(it => list.appendChild(card(it)));
  const dw = $("#droppedWrap");
  dw.innerHTML = "";
  if (state.dropped && state.dropped.length) {
    dw.appendChild(el("h2", null, "No longer flagged since the last run"));
    dw.appendChild(el("div", "dropped", state.dropped.join("  \u00b7  ")));
  }
}

async function load() {
  try { state = await (await fetch("/api/briefing")).json(); } catch (e) {}
  render();
}

$("#switch").onclick = async () => {
  const d = await post("/api/switch", {});
  state.ai_enabled = d.ai_enabled;
  render();
};

$("#refresh").onclick = async () => {
  const b = $("#refresh");
  b.disabled = true;
  b.textContent = "Running pipeline... (can take a minute)";
  try {
    const d = await post("/api/refresh", {});
    if (!d.ok) alert(d.error || "Refresh failed.");
  } catch (e) { alert("Refresh failed."); }
  b.disabled = false;
  b.innerHTML = "&#8635; Refresh briefing";
  load();
};

async function ask() {
  const q = $("#q").value.trim();
  if (!q) return;
  const out = $("#askout");
  out.innerHTML = "";
  const box = el("div", "askbox", state.ai_enabled ? "Thinking locally..." : "Searching...");
  out.appendChild(box);
  let d;
  try { d = await post("/api/ask", {question: q}); }
  catch (e) { box.textContent = "Could not reach the local server."; return; }
  box.innerHTML = "";
  if (d.answer) {
    box.appendChild(el("div", "label", "AI answer (local model)"));
    box.appendChild(el("div", "text", d.answer));
  } else if (!d.ai) {
    box.appendChild(el("div", "label", "AI is off - keyword search only"));
  }
  if (d.error) box.appendChild(el("div", "err", d.error));
  if (d.matches && d.matches.length) {
    box.appendChild(el("div", "label", "Related items"));
    d.matches.forEach(m => box.appendChild(
      el("div", "match", "[" + m.priority.toUpperCase() + "] " + m.title + ": " + m.detail)));
  } else if (!d.answer) {
    box.appendChild(el("div", "text", "No matching items found."));
  }
}
$("#askbtn").onclick = ask;
$("#q").addEventListener("keydown", e => { if (e.key === "Enter") ask(); });

$("#upload").onclick = () => $("#file").click();
$("#file").onchange = async () => {
  const f = $("#file").files[0];
  if (!f) return;
  const b = $("#upload");
  b.disabled = true;
  b.textContent = "Reading your file...";
  const out = $("#uploadout");
  out.innerHTML = "";
  const box = el("div", "askbox");
  try {
    const fd = new FormData();
    fd.append("file", f);
    const d = await (await fetch("/api/upload", {method: "POST", body: fd})).json();
    if (d.ok) {
      box.appendChild(el("div", "label", "How LedgerMind read your file"));
      d.report.forEach(line => box.appendChild(el("div", "match", line)));
    } else {
      box.appendChild(el("div", "err", d.error || "Upload failed."));
    }
  } catch (e) { box.appendChild(el("div", "err", "Upload failed.")); }
  out.appendChild(box);
  b.disabled = false;
  b.innerHTML = "&#8593; Upload your data";
  $("#file").value = "";
  load();
};

load();
</script>
</body>
</html>
"""


@app.get("/")
def index():
    return Response(PAGE, mimetype="text/html")


def _ollama_status():
    try:
        requests.get(OLLAMA_BASE, timeout=2)
        return "Ollama is running."
    except Exception:  # noqa: BLE001
        return "WARNING: Ollama is not reachable. Start it (ollama serve) or the AI parts will show an error."


def _lan_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("10.255.255.255", 1))  # no traffic is sent; it just picks the right network card
        return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"
    finally:
        s.close()


if __name__ == "__main__":
    url = f"http://127.0.0.1:{PORT}"
    print("\nLedgerMind is starting...")
    print(_ollama_status())
    print(f"\nOpen your briefing: {url}")
    if HOST not in ("127.0.0.1", "localhost"):
        if not TOKEN:
            TOKEN = secrets.token_urlsafe(8)
        print(f"Other devices on the same Wi-Fi: http://{_lan_ip()}:{PORT}/?token={TOKEN}")
        print("(Anyone with that link can see your briefing. Only share it on a network you trust.)")
    print("(Press Ctrl+C in this window to stop)\n")
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host=HOST, port=PORT, debug=False)