import json
import os

MEMORY_PATH = "memory/memory.json"

def load_memory():
    if not os.path.exists(MEMORY_PATH):
        return {"seen_anomalies": [], "preferences": {"unusual_outflow_threshold": 300000}}
    with open(MEMORY_PATH) as f:
        return json.load(f)

def save_memory(memory):
    os.makedirs("memory", exist_ok=True)
    with open(MEMORY_PATH, "w") as f:
        json.dump(memory, f, indent=2)