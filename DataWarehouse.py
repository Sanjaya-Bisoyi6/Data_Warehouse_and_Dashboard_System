import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sqlalchemy import create_engine
import urllib
import datetime
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ==================================================
# 🔹 1. SQL Server Connection
# ==================================================
params = urllib.parse.quote_plus(
    "DRIVER={SQL Server};"
    "SERVER=SANJAYA\\SQLEXPRESS;"
    "DATABASE=DataWarehouse;"
    "Trusted_Connection=yes;"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")
print("\n🔗 Connecting to SQL Server...")
print("✅ Connected successfully to DataWarehouse!\n")

# ==================================================
# 🔹 2. Fetch All Tables
# ==================================================
tables_query = """
SELECT TABLE_SCHEMA, TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_TYPE = 'BASE TABLE';
"""
tables_df = pd.read_sql(tables_query, engine)
print("📋 Tables Found in DataWarehouse:")
print(tables_df, "\n")

# ==================================================
# 🔹 3. Initialize Results
# ==================================================
results = []

# ==================================================
# 🔹 4. Table-by-Table Full Data Quality Check
# ==================================================
for _, row in tables_df.iterrows():
    schema = row["TABLE_SCHEMA"]
    table_name = row["TABLE_NAME"]
    full_name = f"[{schema}].[{table_name}]"
    print(f"\n🔍 Testing full table: {full_name}")

    try:
        # Load full table
        df = pd.read_sql(f"SELECT * FROM {full_name}", engine)
        total_rows, total_cols = df.shape

        if total_rows == 0:
            results.append({
                "Table": full_name,
                "Rows": 0,
                "Columns": total_cols,
                "Null_Values": 0,
                "Duplicates": 0,
                "Whitespace_Cells": 0,
                "Invalid_Sales": 0,
                "Invalid_Dates": 0,
                "Integrity_Score": 0,
                "Status": "❌ Empty Table"
            })
            print(f"⚠ {full_name} is empty.")
            continue

        # --- Nulls ---
        null_count = df.isnull().sum().sum()
        null_ratio = round((null_count / (total_rows * total_cols)) * 100, 2)

        # --- Duplicates ---
        duplicate_count = df.duplicated().sum()
        dup_ratio = round((duplicate_count / total_rows) * 100, 2)

        # --- Whitespace ---
        whitespace_count = 0
        for col in df.select_dtypes(include=['object']).columns:
            whitespace_count += df[col].apply(lambda x: isinstance(x, str) and x.strip() != x).sum()
        white_ratio = round((whitespace_count / (total_rows * total_cols)) * 100, 2)

        # --- Invalid Sales / Dates ---
        invalid_sales = 0
        invalid_dates = 0
        for col in df.columns:
            col_lower = col.lower()
            if "sales" in col_lower or "amount" in col_lower:
                invalid_sales += (df[col] < 0).sum()
            if "date" in col_lower:
                invalid_dates += df[col].apply(
                    lambda x: isinstance(x, datetime.datetime)
                    and (x.year < 2000 or x.year > datetime.datetime.now().year)
                ).sum()

        # --- Integrity Score ---
        integrity_score = max(0, 100 - (null_ratio + dup_ratio + white_ratio + invalid_sales + invalid_dates))

        # --- Status ---
        if integrity_score >= 95:
            status = "✅ Excellent"
        elif integrity_score >= 80:
            status = "🟡 Moderate"
        else:
            status = "🔴 Poor"

        # --- Append result ---
        results.append({
            "Table": full_name,
            "Rows": total_rows,
            "Columns": total_cols,
            "Null_Values": null_count,
            "Duplicates": duplicate_count,
            "Whitespace_Cells": whitespace_count,
            "Invalid_Sales": invalid_sales,
            "Invalid_Dates": invalid_dates,
            "Integrity_Score": round(integrity_score, 2),
            "Status": status
        })

        print(f"✅ Tested {full_name} successfully.")
        print(f"Rows: {total_rows}, Columns: {total_cols}")
        print(f"Nulls: {null_count}, Duplicates: {duplicate_count}, Whitespace: {whitespace_count}")
        print(f"Integrity: {round(integrity_score,2)}% → {status}")

    except MemoryError:
        print(f"⚠ Table {full_name} too large, testing sample instead.")
        sample_df = pd.read_sql(f"SELECT TOP 100000 * FROM {full_name}", engine)
        total_rows, total_cols = sample_df.shape
        null_count = sample_df.isnull().sum().sum()
        duplicate_count = sample_df.duplicated().sum()
        whitespace_count = 0
        for col in sample_df.select_dtypes(include=['object']).columns:
            whitespace_count += sample_df[col].apply(lambda x: isinstance(x, str) and x.strip() != x).sum()

        results.append({
            "Table": full_name,
            "Rows": total_rows,
            "Columns": total_cols,
            "Null_Values": null_count,
            "Duplicates": duplicate_count,
            "Whitespace_Cells": whitespace_count,
            "Integrity_Score": "Sample Tested",
            "Status": "⚠ Large Table - Partial Test"
        })

    except Exception as e:
        results.append({
            "Table": full_name,
            "Status": f"❌ Error: {e}"
        })
        print(f"❌ Error testing {full_name}: {e}")

# ==================================================
# 🔹 5. Save Reports
# ==================================================
result_df = pd.DataFrame(results)
output_dir = os.path.join(os.getcwd(), "Testing_Reports")
os.makedirs(output_dir, exist_ok=True)

excel_path = os.path.join(output_dir, "Retail_Data_Test_Report.xlsx")
csv_path = os.path.join(output_dir, "Retail_Data_Test_Report.csv")

# Save both Excel & CSV (for Power BI)
result_df.to_excel(excel_path, index=False)
result_df.to_csv(csv_path, index=False)

# Format Excel
wb = load_workbook(excel_path)
ws = wb.active
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True)
for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")

for row in ws.iter_rows(min_row=2, max_col=ws.max_column):
    status = row[-1].value
    if "Excellent" in str(status):
        color = "C6EFCE"
    elif "Moderate" in str(status):
        color = "FFF2CC"
    elif "Poor" in str(status):
        color = "FFC7CE"
    else:
        color = "FFFFFF"
    for cell in row:
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

wb.save(excel_path)
print(f"\n📊 Excel Report: {excel_path}")
print(f"📈 Power BI CSV Report: {csv_path}")

# ==================================================
# 🔹 6. Visualization
# ==================================================
try:
    sns.set(style="whitegrid")

    # Bar Chart
    plt.figure(figsize=(10, 6))
    sns.barplot(x="Table", y="Integrity_Score", hue="Status", data=result_df)
    plt.title("🧭 Data Integrity Score by Table (Full Data)")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "Integrity_BarChart.png"))
    plt.show()

    # Pie Chart
    plt.figure(figsize=(6, 6))
    result_df["Status"].value_counts().plot.pie(
        autopct="%1.1f%%", startangle=90, colors=["#66BB6A", "#FFD54F", "#EF5350"]
    )
    plt.title("📊 Overall Data Quality Status Distribution")
    plt.ylabel("")
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "Integrity_PieChart.png"))
    plt.show()

    print("✅ Visuals generated successfully!")

except Exception as e:
    print(f"⚠ Visualization Error: {e}")

print("\n✅ Full-Table Data Testing Completed Successfully!")