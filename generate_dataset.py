"""
Synthetic dataset generator: Invoice Reconciliation & AR Aging Automation
Produces invoices_raw.csv and payments_raw.csv with deliberate, documented
messiness so the reconciliation/data-quality VBA logic has real issues to catch.
Seeded for reproducibility.
"""
import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

random.seed(42)
np.random.seed(42)

# ---------------- Customer pool ----------------
prefixes = ["Summit","Meridian","Falcon","Harbor","Atlas","Vantage","Crescent","Ironwood",
            "Bluepeak","Northgate","Cedar","Anchor","Pinnacle","Silverline","Redwood",
            "Brightpath","Sterling","Union","Cascade","Granite"]
suffixes = ["Logistics","Supply Co.","Group","Partners","Solutions","Manufacturing",
            "Retail","Holdings","Industries","Consulting"]
customers = sorted(set(f"{p} {s}" for p in prefixes for s in suffixes))
random.shuffle(customers)
customers = customers[:150]

# ---------------- Config ----------------
N_BASE_INVOICES = 3000
N_DUPLICATES = 18
FY_START = datetime(2025, 1, 1)
FY_END = datetime(2025, 12, 31)
terms_options = [15, 30, 30, 30, 45, 60]  # weighted toward Net 30
methods = ["Bank Transfer", "Credit Card", "Check", "ACH"]

def random_date(start, end):
    return start + timedelta(days=random.randint(0, (end - start).days))

# ---------------- Base invoices ----------------
invoices = []
for i in range(N_BASE_INVOICES):
    inv_date = random_date(FY_START, FY_END)
    terms = random.choice(terms_options)
    amount = round(np.random.lognormal(mean=7.3, sigma=0.9), 2)
    amount = max(150, min(amount, 18000))
    invoices.append({
        "Invoice ID": f"INV-{10001+i}",
        "Customer Name": random.choice(customers),
        "Invoice Date": inv_date,
        "Due Date": inv_date + timedelta(days=terms),
        "Amount": amount,
        "Region": random.choice(["North America", "EMEA", "APAC", "LATAM"])
    })
inv_df = pd.DataFrame(invoices)
inv_df["_status"] = np.random.choice(
    ["Paid", "Partial", "Unpaid", "Overpaid"], size=N_BASE_INVOICES, p=[0.62, 0.15, 0.18, 0.05]
)

# ---------------- Payments ----------------
payments = []
pay_counter = 50001
for _, row in inv_df.iterrows():
    status, inv_id = row["_status"], row["Invoice ID"]
    inv_date, due_date, amount = row["Invoice Date"], row["Due Date"], row["Amount"]
    span = max((due_date - inv_date).days, 1)

    if status == "Unpaid":
        continue
    elif status == "Paid":
        pay_date = inv_date + timedelta(days=random.randint(1, span + 20))
        payments.append({"Payment ID": f"PAY-{pay_counter}", "Invoice ID": inv_id,
                          "Payment Date": pay_date, "Amount Paid": amount,
                          "Payment Method": random.choice(methods)})
        pay_counter += 1
    elif status == "Partial":
        n_splits = random.choice([2, 3])
        remaining = amount * random.uniform(0.4, 0.85)
        pts = np.sort(np.random.uniform(0, remaining, n_splits - 1))
        splits = np.diff(np.concatenate(([0], pts, [remaining])))
        for s in splits:
            pay_date = inv_date + timedelta(days=random.randint(1, span + 40))
            payments.append({"Payment ID": f"PAY-{pay_counter}", "Invoice ID": inv_id,
                              "Payment Date": pay_date, "Amount Paid": round(float(s), 2),
                              "Payment Method": random.choice(methods)})
            pay_counter += 1
    elif status == "Overpaid":
        pd1 = inv_date + timedelta(days=random.randint(1, span + 10))
        pd2 = pd1 + timedelta(days=random.randint(1, 15))
        payments.append({"Payment ID": f"PAY-{pay_counter}", "Invoice ID": inv_id,
                          "Payment Date": pd1, "Amount Paid": amount,
                          "Payment Method": random.choice(methods)})
        pay_counter += 1
        payments.append({"Payment ID": f"PAY-{pay_counter}", "Invoice ID": inv_id,
                          "Payment Date": pd2, "Amount Paid": round(amount * random.uniform(0.1, 0.3), 2),
                          "Payment Method": random.choice(methods)})
        pay_counter += 1
pay_df = pd.DataFrame(payments)

# ---------------- Inject: duplicate invoice rows ----------------
dup_idx = np.random.choice(inv_df.index, size=N_DUPLICATES, replace=False)
dup_rows = inv_df.loc[dup_idx].copy()
dup_rows["Amount"] = (dup_rows["Amount"] * np.random.uniform(0.95, 1.05, size=len(dup_rows))).round(2)
inv_df_final = pd.concat([inv_df, dup_rows], ignore_index=True)

# ---------------- Inject: orphaned payments (Invoice ID with no matching invoice) ----------------
n_orphaned = int(len(pay_df) * 0.02)
orphan_rows = []
for _ in range(n_orphaned):
    fake_id = f"INV-{random.randint(19000, 19999)}"
    orphan_rows.append({"Payment ID": f"PAY-{pay_counter}", "Invoice ID": fake_id,
                         "Payment Date": random_date(FY_START, FY_END),
                         "Amount Paid": round(random.uniform(200, 5000), 2),
                         "Payment Method": random.choice(methods)})
    pay_counter += 1
pay_df = pd.concat([pay_df, pd.DataFrame(orphan_rows)], ignore_index=True)

# ---------------- Inject: blanks ----------------
blank_cust_idx = np.random.choice(inv_df_final.index, size=int(len(inv_df_final)*0.01), replace=False)
inv_df_final.loc[blank_cust_idx, "Customer Name"] = ""
blank_region_idx = np.random.choice(inv_df_final.index, size=int(len(inv_df_final)*0.015), replace=False)
inv_df_final.loc[blank_region_idx, "Region"] = ""

# ---------------- Inject: whitespace noise on Invoice ID ----------------
ws_idx = np.random.choice(inv_df_final.index, size=int(len(inv_df_final)*0.015), replace=False)
inv_df_final.loc[ws_idx, "Invoice ID"] = inv_df_final.loc[ws_idx, "Invoice ID"].apply(
    lambda x: f" {x} " if random.random() < 0.5 else f"{x} "
)

# ---------------- Format dates (invoices = consistent ISO, payments = mixed) ----------------
inv_df_final["Invoice Date"] = pd.to_datetime(inv_df_final["Invoice Date"]).dt.strftime("%Y-%m-%d")
inv_df_final["Due Date"] = pd.to_datetime(inv_df_final["Due Date"]).dt.strftime("%Y-%m-%d")

pay_df["Payment Date"] = pd.to_datetime(pay_df["Payment Date"])
mixed_mask = np.random.rand(len(pay_df)) < 0.12
pay_df["Payment Date"] = np.where(
    mixed_mask,
    pay_df["Payment Date"].dt.strftime("%m/%d/%Y"),
    pay_df["Payment Date"].dt.strftime("%Y-%m-%d")
)

# ---------------- Shuffle + save ----------------
status_counts = inv_df["_status"].value_counts().to_dict()
inv_df_final = inv_df_final.drop(columns=["_status"]).sample(frac=1, random_state=42).reset_index(drop=True)
pay_df = pay_df.sample(frac=1, random_state=42).reset_index(drop=True)

inv_df_final.to_csv("invoices_raw.csv", index=False)
pay_df.to_csv("payments_raw.csv", index=False)

print("=== GENERATION REPORT ===")
print(f"invoices_raw.csv rows: {len(inv_df_final)}  (base {N_BASE_INVOICES} + {N_DUPLICATES} injected duplicates)")
print(f"payments_raw.csv rows: {len(pay_df)}")
print(f"Status mix (of {N_BASE_INVOICES} base invoices): {status_counts}")
print(f"Orphaned payments (Invoice ID matches nothing): {n_orphaned}")
print(f"Blank Customer Name: {len(blank_cust_idx)}")
print(f"Blank Region: {len(blank_region_idx)}")
print(f"Whitespace-noised Invoice IDs: {len(ws_idx)}")
print(f"Payments with MM/DD/YYYY format (rest are YYYY-MM-DD): {int(mixed_mask.sum())}")
print("\n--- invoices_raw.csv sample ---")
print(inv_df_final.head(5).to_string())
print("\n--- payments_raw.csv sample ---")
print(pay_df.head(5).to_string())
