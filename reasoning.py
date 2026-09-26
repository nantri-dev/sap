import requests
import json
import chromadb
from sentence_transformers import SentenceTransformer
from memory_store import load_memory, save_memory

model = SentenceTransformer("all-MiniLM-L6-v2")
client = chromadb.PersistentClient(path="./chroma_db")
collection = client.get_or_create_collection("ledgermind_knowledge")

OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "llama3.2:1b"

def retrieve_context(query, top_k=8):
    query_emb = model.encode([query]).tolist()
    results = collection.query(query_embeddings=query_emb, n_results=top_k)
    return results["documents"][0]

def build_prompt(context, memory):
    context_text = "\n".join(f"- {c}" for c in context)
    threshold = memory["preferences"]["unusual_outflow_threshold"]
    return f"""You are LedgerMind, a private finance briefing assistant. All data stays local.

Here is today's relevant financial data:
{context_text}

The user considers any outflow above {threshold} as unusual and worth flagging.
Previously flagged anomalies (don't repeat unless status changed): {memory['seen_anomalies']}

Generate a short morning briefing as a JSON array. Each item must have:
"category" ("summary", "alert", or "action"), "title", "detail", "priority" ("low","medium","high").
Only output the JSON array, nothing else.
"""

def call_ollama(prompt):
    response = requests.post(OLLAMA_URL, json={
        "model": OLLAMA_MODEL,
        "prompt": prompt,
        "stream": False
    })
    return response.json()["response"]

def parse_briefing(raw_text):
    try:
        start = raw_text.find("[")
        end = raw_text.rfind("]") + 1
        return json.loads(raw_text[start:end])
    except Exception:
        return [{"category": "summary", "title": "Raw model output",
                  "detail": raw_text, "priority": "low"}]

def generate_briefing():
    memory = load_memory()
    context = retrieve_context("Give today's financial briefing: cash position, loans due, anomalies")
    prompt = build_prompt(context, memory)
    raw = call_ollama(prompt)
    briefing = parse_briefing(raw)

    for item in briefing:
        if item["category"] == "alert" and item["title"] not in memory["seen_anomalies"]:
            memory["seen_anomalies"].append(item["title"])
    save_memory(memory)

    return briefing

if __name__ == "__main__":
    result = generate_briefing()
    print(json.dumps(result, indent=2))