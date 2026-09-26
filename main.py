import pandas as pd
from ingest import ingest
from reasoning import generate_briefing

def run():
    print("Step 1: Ingesting Excel data into knowledge base...")
    ingest()

    print("Step 2: Running local reasoning...")
    briefing = generate_briefing()

    print("Step 3: Writing output for Power BI...")
    df = pd.DataFrame(briefing)
    df.to_csv("output/briefing.csv", index=False)
    print("Done. output/briefing.csv is ready for Power BI.")

if __name__ == "__main__":
    run()