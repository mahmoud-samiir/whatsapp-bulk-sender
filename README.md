# 📱 WhatsApp Bulk Message Sender (Python + Selenium)

This project is a tool for sending bulk WhatsApp messages using **WhatsApp Web** and **Selenium**.  
It reads phone numbers from an Excel file and sends a custom message to each number with a **1-minute delay** between messages.

---

## 🚀 Features
- Read phone numbers from Excel files (.xlsx or .xls).
- Supports columns named `Phone`, `Number`, or defaults to the first column.
- Cleans phone numbers (removes spaces and symbols).
- Logs sending results (success / failed) into a log file.
- Runs automatically on **Microsoft Edge (WebDriver)**.

---

## 🛠️ Requirements
Make sure you have:
- Python 3.8 or later
- Microsoft Edge browser
- Microsoft Edge WebDriver (matching your browser version)

Install the required libraries:
```bash
pip install selenium pandas openpyxl
