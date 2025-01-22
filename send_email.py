import smtplib
import pandas as pd
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import time

# Load configuration from config.json
def load_config(config_path="config.json"):
    with open(config_path, "r") as file:
        return json.load(file)

# Load student grades from CSV or Excel file
def load_grades(file_path):
    file_extension = os.path.splitext(file_path)[-1].lower()
    
    if file_extension == ".csv":
        return pd.read_csv(file_path)
    elif file_extension in [".xls", ".xlsx"]:
        return pd.read_excel(file_path)
    else:
        raise ValueError("Unsupported file format. Please use CSV or Excel.")

# Identify relevant columns dynamically
def identify_columns(df):
    possible_name_columns = ["name", "student name", "full name"]
    possible_email_columns = ["email", "student email", "contact email"]
    
    # Convert column names to lowercase for case-insensitive matching
    df.columns = [col.lower() for col in df.columns]
    
    name_col = next((col for col in possible_name_columns if col in df.columns), None)
    email_col = next((col for col in possible_email_columns if col in df.columns), None)
    
    if not name_col or not email_col:
        raise ValueError("Could not identify 'Name' or 'Email' columns in the dataset.")
    
    return name_col, email_col

# Generate a personalized email message
def generate_email(sender_email, student_name, student_email, grades):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = student_email
    msg['Subject'] = "Your Assignment Grades"

    # Format grades into a message
    grade_details = "\n".join([f"{col}: {grade}" for col, grade in grades.items()])
    body = f"""
    Dear {student_name},

    Here are your assignment grades:
    {grade_details}

    Best,
    Gane Ka-Shu Wong
    """

    msg.attach(MIMEText(body, 'plain'))
    return msg

# Send email
def send_email(config, msg, student_email):
    try:
        server = smtplib.SMTP(config["SMTP_SERVER"], config["SMTP_PORT"])
        server.starttls()  # Secure connection
        server.login(config["SENDER_EMAIL"], config["SENDER_PASSWORD"])
        server.sendmail(config["SENDER_EMAIL"], student_email, msg.as_string())
        server.quit()
        print(f"Email sent to {student_email}")
    except Exception as e:
        print(f"Failed to send email to {student_email}: {e}")

# Main function
def main():
    config = load_config("config.json")  # Load SMTP credentials
    file_path = "grades.xlsx"  # Replace with your actual file
    df = load_grades(file_path)
    
    name_col, email_col = identify_columns(df)

    for _, row in df.iterrows():
        time.sleep(0.5)
        student_name = row[name_col]
        student_email = row[email_col]
        grades = row.drop([name_col, email_col]).to_dict()  # Exclude name & email

        msg = generate_email(config["SENDER_EMAIL"], student_name, student_email, grades)
        send_email(config, msg, student_email)

if __name__ == "__main__":
    main()
