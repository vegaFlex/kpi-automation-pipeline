# KPI Automation Pipeline

## Overview
This project demonstrates an end-to-end automated KPI reporting pipeline using Python, PostgreSQL, and Power BI.

The system extracts data from a PostgreSQL database, validates and processes it in Python, generates structured Excel reports, and sends them automatically via email.

A Power BI dashboard is built directly on top of PostgreSQL, enabling real-time KPI tracking without relying on static Excel files.

---

## Dashboard Preview

![KPI Automation Pipeline Cover](cover%20photo.png)

### KPI Overview
![KPI](images/kpi_overview.png)

Displays total revenue and the top-performing manager based on actual revenue.

### Manager Performance
![Performance](images/manager_performance.png)

Compares revenue across managers to quickly identify top and low performers.

### Detailed Metrics
![Metrics](images/detailed_metrics.png)

Shows detailed KPIs including revenue, completed orders, and achievement percentage calculated from base metrics.

---

## Key Features
- SQL-based data extraction from PostgreSQL
- Data validation and type handling in Python
- Automated Excel report generation with formatting
- Logging system for monitoring pipeline execution
- Automated email delivery with report attachment
- Scheduled daily execution via Windows Task Scheduler
- Power BI dashboard connected directly to PostgreSQL (no Excel dependency)

---

## Tech Stack
- Python (pandas, sqlalchemy)
- PostgreSQL
- openpyxl
- smtplib
- python-dotenv
- Power BI

---

## Project Structure
- `main.py` – main pipeline logic
- `.env` – environment variables (excluded from repo)
- `requirements.txt` – project dependencies
- `run_kpi_report.bat` – automation entry point
- `reports/` – generated reports (ignored in Git)
- `images/` – dashboard screenshots

---

## How It Works
1. Extracts transactional data from PostgreSQL using SQL
2. Validates and cleans the dataset in Python
3. Aggregates KPI metrics (revenue, orders, performance)
4. Generates formatted Excel report
5. Sends report automatically via email
6. Logs execution for monitoring and debugging
7. Power BI connects directly to PostgreSQL for visualization

---

## Automation
The pipeline runs automatically on a daily schedule using Windows Task Scheduler, eliminating manual reporting.

---

## Data Modeling Insight
Achievement percentage is not aggregated directly.
Instead, it is recalculated in Power BI using:

Revenue / Target

This ensures correct KPI calculation and avoids common aggregation mistakes.

---

## Security
- Credentials are stored in `.env`
- Sensitive data is excluded via `.gitignore`
- No secrets are exposed in the repository

---

## Business Value
- Eliminates manual Excel reporting
- Reduces human error
- Provides consistent KPI tracking
- Enables faster, data-driven decision-making
- Bridges backend automation with business intelligence

---

## Result
A production-style data pipeline combining:
- data extraction (SQL)
- processing (Python)
- automation (scheduler + email)
- visualization (Power BI)

This reflects a real-world business reporting workflow.
