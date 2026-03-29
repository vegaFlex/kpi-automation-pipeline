# KPI Automation Pipeline

## Overview
This project automates the generation and delivery of KPI reports using Python and PostgreSQL.

It eliminates manual reporting by extracting data from a database, validating it, generating a formatted Excel report, and sending it automatically via email.

## Features
- SQL data extraction from PostgreSQL
- Data validation and cleaning
- Automated Excel report generation
- Excel formatting (headers, column widths)
- Logging system for monitoring
- Automated email delivery
- Scheduled execution via Windows Task Scheduler

## Tech Stack
- Python (pandas, sqlalchemy)
- PostgreSQL
- openpyxl
- smtplib
- dotenv

## Project Structure
- `main.py` – main pipeline logic
- `.env` – environment variables (not included in repo)
- `requirements.txt` – dependencies
- `run_kpi_report.bat` – automation script
- `reports/` – generated reports (ignored in Git)

## How It Works
1. Extracts data from PostgreSQL
2. Validates and cleans the dataset
3. Generates KPI report in Excel
4. Saves the file with dynamic timestamp
5. Sends report via email
6. Logs all actions for monitoring

## Automation
The script is scheduled using Windows Task Scheduler to run daily without manual intervention.

## Security
Sensitive data such as database credentials and email passwords are stored in a `.env` file and are not included in the repository.

## Result
Fully automated reporting system that replaces manual Excel work and reduces human error.