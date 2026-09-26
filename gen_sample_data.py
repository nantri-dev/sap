import pandas as pd

data = [
    {"date": "2026-09-20", "type": "loan", "counterparty": "Deutsche Bank", "amount": 2500000, "account": "Main Operating", "due_date": "2026-09-28"},
    {"date": "2026-09-18", "type": "deposit", "counterparty": "HSBC", "amount": 1100000, "account": "FX Reserve", "due_date": "2026-10-01"},
    {"date": "2026-09-15", "type": "revolving_credit", "counterparty": "Citi", "amount": 5000000, "account": "Main Operating", "due_date": "2026-10-03"},
    {"date": "2026-09-25", "type": "outflow", "counterparty": "Subsidiary EU", "amount": 430000, "account": "EU Subsidiary", "due_date": ""},
    {"date": "2026-09-25", "type": "deal", "counterparty": "Barclays", "amount": 750000, "account": "Main Operating", "due_date": ""},
    {"date": "2026-09-25", "type": "deal", "counterparty": "Barclays", "amount": 750000, "account": "Main Operating", "due_date": ""},
    {"date": "2026-09-26", "type": "cash_balance", "counterparty": "Internal", "amount": 18400000, "account": "Main Operating", "due_date": ""},
    {"date": "2026-09-25", "type": "cash_balance_prior", "counterparty": "Internal", "amount": 18020000, "account": "Main Operating", "due_date": ""},
]

df = pd.DataFrame(data)
df.to_excel("data/sample_data.xlsx", index=False)
print("Created data/sample_data.xlsx")