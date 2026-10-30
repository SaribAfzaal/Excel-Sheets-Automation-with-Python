from pathlib import Path

import pandas as pd

out = Path(__file__).parent / "sample_input"
out.mkdir(exist_ok=True)

pd.DataFrame({
    "Order ID": [1001, 1002, 1003, 1004, 1005],
    "Customer": ["Ahmed Khan", "Fatima Malik", "Bilal Ahmed", "Ayesha Siddiqui", "Hamza Qureshi"],
    "Region": ["Punjab", "Sindh", "Khyber Pakhtunkhwa", "Balochistan", "Punjab"],
    "Sales": [1500, 820, 2300, None, 640],
    "Quantity": [15, 8, 23, 5, 6],
    "Discount": [0.05, 0.05, 0.10, 0.0, 0.0],
}).to_excel(out / "sales_q1.xlsx", index=False)

pd.DataFrame({
    "customer": ["Fatima Malik", "Usman Sheikh", "Zainab Raza", "Imran Chaudhry"],
    "order id": [1002, 1006, 1007, 1008],
    "SALES": [820, 1975, 3100, 410],
    "region": ["Sindh", "Khyber Pakhtunkhwa", None, "Balochistan"],
    "Quantity": [8, 19, 31, 4],
    "Discount": [0.05, 0.10, 0.15, 0.0],
}).to_csv(out / "sales_q2.csv", index=False)

with pd.ExcelWriter(out / "sales_regional.xlsx") as xw:
    pd.DataFrame({
        "Order ID": [1009, 1010, 1011],
        "Customer": ["Sana Iqbal", "Tariq Mehmood", None],
        "Region": ["Punjab", "Sindh", "Khyber Pakhtunkhwa"],
        "Sales": [980, 1200, 2750],
    }).to_excel(xw, sheet_name="Jan", index=False)
    pd.DataFrame({
        "Order ID": [1012, 1013],
        "Customer": ["Hina Butt", "Kamran Akhtar"],
        "Region": ["Balochistan", "Punjab"],
        "Sales": ["$1,100", "1,450"],
        "Quantity": [11, 14],
    }).to_excel(xw, sheet_name="Feb", index=False)

print(f"Sample data written to {out}")
