import smtplib
import pandas as pd
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import time
import argparse


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
    
    # Check for 'first name' and 'last name' columns
    if 'first name' in df.columns and 'last name' in df.columns:
        # Combine 'first name' and 'last name' into a single 'name' column
        df['name'] = df['first name'].str.strip() + ' ' + df['last name'].str.strip()
        # Remove the original 'first name' and 'last name' columns
        df.drop(['first name', 'last name'], axis=1, inplace=True)
        name_col = 'name'
    else:
        # Check for existing name columns
        name_col = next((col for col in possible_name_columns if col in df.columns), None)
    
    # Determine the email column
    if 'ccid' in df.columns:
        email_col = 'ccid'
        df['ccid']=df['ccid']+'@ualberta.ca'
    else:
        email_col = next((col for col in possible_email_columns if col in df.columns), None)
    
    if not name_col or not email_col:
        raise ValueError("Could not identify 'Name' or 'Email' columns in the dataset.")
    if 'id' in df.columns:
      id_col='id'
    else:
      id_col=None
    return name_col, email_col, id_col


# Generate a personalized email message
def generate_email(sender_email, sender_name, student_name, student_email, student_id, grades, subject= None):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = student_email
    if subject:
      msg['Subject'] = subject
    else:
      msg['Subject'] = "Your Assignment Grades"

    # Format grades into a message
    if student_id:
      id_message= f' with student ID: {student_id}'
    else:
      id_message=''
    grade_details = "\n".join([f"{col.capitalize()}: {grade}" for col, grade in grades.items()])
    body = f"""
Dear {student_name}{id_message},

Here are your assignment grades:
{grade_details}

Best,
{sender_name}
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
    parser = argparse.ArgumentParser(description="Process grades file address.")
    
    parser.add_argument(
        "grades_address",
        type=str,
        nargs="?",  # Makes it optional
        default="grades.xlsx",  # Default value if not provided
        help="Path to the grades file (default: grades.xlsx)"
    )
    parser.add_argument(
        "--config",
        type=str,
        default="config.json",
        help="Path to the configuration file (default: config.json)"
    )

    args = parser.parse_args()
    config = load_config(args.config)  # Load SMTP credentials
    file_path = args.grades_address  # Replace with your actual file
    df = load_grades(file_path)
    name_col, email_col, id_col = identify_columns(df)
    assignments_name=file_path.split('/')[1].split('.')[0]
    subject = f'Grades for {assignments_name}'
    for _, row in df.iterrows():
        time.sleep(0.5)
        
        student_name = row[name_col]
        student_email = row[email_col]
        if id_col:
          student_id= row[id_col]
          grades = row.drop([name_col, email_col, id_col]).to_dict()  # Exclude name & email
          msg = generate_email(config["SENDER_EMAIL"],config["SENDER_NAME"], student_name, student_email, student_id, grades, subject)
        else:
          grades = row.drop([name_col, email_col]).to_dict()  # Exclude name & email
          msg = generate_email(config["SENDER_EMAIL"],config["SENDER_NAME"], student_name, student_email, None, grades, subject)
        send_email(config, msg, student_email)

if __name__ == "__main__":
    main()
