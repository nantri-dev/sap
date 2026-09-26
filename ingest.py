import pandas as pd
import chromadb
from sentence_transformers import SentenceTransformer

model = SentenceTransformer("all-MiniLM-L6-v2")  # small, runs locally, no API calls
client = chromadb.PersistentClient(path="./chroma_db")
collection = client.get_or_create_collection("ledgermind_knowledge")

def row_to_text(row):
    return (f"On {row['date']}, a {row['type']} transaction of {row['amount']} "
            f"with {row['counterparty']} on account {row['account']}"
            + (f", due {row['due_date']}" if row['due_date'] else "") + ".")

def ingest():
    df = pd.read_excel("data/sample_data.xlsx")
    docs = [row_to_text(r) for _, r in df.iterrows()]
    ids = [f"row-{i}" for i in range(len(docs))]
    embeddings = model.encode(docs).tolist()

    # clear old entries so re-running doesn't duplicate
    existing = collection.get()["ids"]
    if existing:
        collection.delete(ids=existing)

    collection.add(documents=docs, embeddings=embeddings, ids=ids)
    print(f"Ingested {len(docs)} records into knowledge base.")

if __name__ == "__main__":
    ingest()