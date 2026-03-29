import os
import pandas as pd
import logging

import smtplib
from email.message import EmailMessage

from sqlalchemy import create_engine

from datetime import datetime
from pathlib import Path

from openpyxl.styles import Font 

from dotenv import load_dotenv
load_dotenv()

logging.basicConfig(
    filename="pipeline.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logging.info("=== PIPELINE START ===")

DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")
DB_HOST = os.getenv("DB_HOST")
DB_PORT = os.getenv("DB_PORT")
DB_NAME = os.getenv("DB_NAME")

EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER")

engine = create_engine(
    f"postgresql+psycopg2://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}"
)

query = """
SELECT
    s.manager_id,
    m.manager_name,
    DATE_TRUNC('month', s.order_date)::date AS month,
    ROUND(SUM(s.amount), 2) AS actual_revenue,
    COUNT(*) AS completed_orders,
    t.target_revenue,
    ROUND(
        (SUM(s.amount) / NULLIF(t.target_revenue, 0)) * 100,
        2
    ) AS achievement_percent
FROM public.sales s
JOIN public.managers m
    ON s.manager_id = m.manager_id
LEFT JOIN public.targets t
    ON t.manager_id = s.manager_id
   AND t.month = DATE_TRUNC('month', s.order_date)::date
WHERE s.status = 'Completed'
  AND s.order_date >= DATE '2026-01-01'
  AND s.order_date < DATE '2026-04-01'
GROUP BY
    s.manager_id,
    m.manager_name,
    DATE_TRUNC('month', s.order_date)::date,
    t.target_revenue
ORDER BY month, actual_revenue DESC;
"""
def validate_data(df):
    # print("=== VALIDATION START ===")
    logging.info("=== VALIDATION START ===")

    # 1. Дата
    df["month"] = pd.to_datetime(df["month"], errors="coerce")

    # 2. Числа
    numeric_cols = ["actual_revenue", "target_revenue", "achievement_percent"]

    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # 3. NULL check
    null_counts = df.isnull().sum()
    # print("NULL values:\n", null_counts)
    logging.info(f"NULL values:\n{null_counts}")

    # 4. Премахни редове с проблеми
    df = df.dropna()

    # 5. Логика проверки
    invalid_revenue = df[df["actual_revenue"] < 0]
    if not invalid_revenue.empty:
        # print("⚠️ Невалидни приходи:", len(invalid_revenue))
        logging.warning(f"Невалидни приходи: {len(invalid_revenue)}")

    # print("=== VALIDATION END ===")
    logging.info("=== VALIDATION END ===")

    return df





df = pd.read_sql(query, engine)
df = validate_data(df)

# df["month"] = df["month"].dt.date
df["month"] = df["month"].dt.strftime("%Y-%m")

print(df.head())
print("Rows:", df.shape[0])


reports_dir = Path("reports")
reports_dir.mkdir(exist_ok=True)

today_str = datetime.today().strftime("%Y-%m-%d")
output_file = reports_dir / f"kpi_report_{today_str}.xlsx"

total_revenue = df["actual_revenue"].sum()
total_orders = df["completed_orders"].sum()

top_manager = df.sort_values(by="actual_revenue", ascending=False).iloc[0]
worst_manager = df.sort_values(by="actual_revenue", ascending=True).iloc[0]

summary_df = pd.DataFrame({
    "Metric": [
        "Total Revenue",
        "Total Orders",
        "Top Manager",
        "Top Manager Revenue",
        "Worst Manager",
        "Worst Manager Revenue"
    ],
    "Value": [
        total_revenue,
        total_orders,
        top_manager["manager_name"],
        top_manager["actual_revenue"],
        worst_manager["manager_name"],
        worst_manager["actual_revenue"]
    ]
})

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    
    df.to_excel(writer, index=False, sheet_name="KPI Report")

    summary_df.to_excel(writer, index=False, sheet_name="Summary")

    workbook = writer.book
    worksheet = writer.sheets["KPI Report"]

    for cell in worksheet[1]:
        # cell.font = cell.font.copy(bold=True)
        cell.font = Font(bold=True)

    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = column_cells[0].column_letter

        for cell in column_cells:
            cell_value = "" if cell.value is None else str(cell.value)
            if len(cell_value) > max_length:
                max_length = len(cell_value)

        worksheet.column_dimensions[column_letter].width = max_length + 2

def send_email(file_path):
    msg = EmailMessage()
    msg["Subject"] = "KPI Report"
    msg["From"] = EMAIL_SENDER
    msg["To"] = EMAIL_RECEIVER

    msg.set_content("Виж прикачения KPI отчет.")

    with open(file_path, "rb") as f:
        file_data = f.read()
        file_name = f.name

    msg.add_attachment(
        file_data,
        maintype="application",
        subtype="octet-stream",
        filename=file_name
    )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
        smtp.send_message(msg)

    logging.info("Email sent successfully")
    print("Email изпратен")


logging.info(f"Report generated: {output_file}")

print(f"Файлът е записан: {output_file}")
send_email(output_file)