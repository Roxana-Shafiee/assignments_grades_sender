# Email Automation Project - Setup Guide

## Overview
This project automates sending personalized emails with students' assignment grades. It reads data from a CSV or Excel file and sends emails to each student based on the provided details.

## Prerequisites
- **Python 3.8 or later**
- Required libraries (install using `requirements.txt`)

## Installation Steps

### 1. Clone the Repository
```bash
git clone https://github.com/Roxana-Shafiee/assignments_grades_sender.git
cd assignments_grades_sender
```

### 2. Set Up a Virtual Environment (Optional but Recommended)
```bash
python -m venv venv
source venv/bin/activate  # On macOS/Linux
venv\Scripts\activate  # On Windows
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Prepare the Configuration File
Create a `config.json` file in the project root with the following structure:
```json
{
    "SMTP_SERVER": "smtp.your-email-provider.com",
    "SMTP_PORT": 587,
    "SENDER_EMAIL": "your-email@example.com",
    "SENDER_PASSWORD": "your-app-password",
    "SENDER_NAME": "your-name"
}
```
> ⚠️ **Do not share this file publicly!**
> ⚠️ **Note: The password is generated for smtp and is not the main password of the email.**

### 5. Prepare the Data File
You need a file (`grades.csv` or `grades.xlsx`) that contains student information with at least the following columns:

| Name         | Email              | Assignment 1 | Assignment 2 | Final Exam |
|-------------|--------------------|--------------|-------------|------------|
| John Doe    | john@example.com   | 85          | 90          | 88         |
| Jane Smith  | jane@example.com   | 92          | 88          | 94         |

- Column names can vary, but the script will try to identify the **Name**, **Email**, and **ID** columns dynamically.
- Additional columns will be treated as assignment grades.

### 6. Run the Script
```bash
python email_grades_sender.py
```
or
```bash
python email_grades_sender.py path/to/your/grades.your_file_format
```
This will:
- Load the configuration and student data.
- Generate personalized emails for each student.
- Send the emails using the SMTP credentials provided.

### 7. Verify Emails Sent
Check your email account's **Sent Items** to verify that the emails were successfully delivered. If errors occur, they will be logged in the console output.

## Troubleshooting
### Common Issues & Fixes
- **Invalid Credentials:** Ensure the email and password in `config.json` are correct.
- **SMTP Errors:** Some providers require enabling "Less Secure Apps" or setting up an **App Password**.
- **Missing Dependencies:** Run `pip install -r requirements.txt` to ensure all required libraries are installed.

## Future Improvements
- Adding support for attachments (e.g., feedback reports)
- Logging email send status to a separate log file
- Customizable email templates

---
