import requests
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from io import BytesIO
from datetime import datetime
import os

# Function to reset the sequence number to 1
def reset_sequence_number(sequence_file):
    with open(sequence_file, 'w') as f:
        f.write("1")
    print("Sequence number reset to 1.")

# Step 1: Obtain the access token
def get_access_token():
    url = "https://accounts.zoho.com/oauth/v2/token"
    payload = {
        'client_id': '1000.BWTPMEU6XB14CRJZLP4Q2X44MEI5JV',
        'client_secret': 'e56ab2e488e480e5cd0245f206186ee274a6bb8a71',
        'grant_type': 'refresh_token',
        'refresh_token': '1000.57424c86722f5ef38b654f08cdea4fad.783f985400136672a72b5c1171a42d8c'
    }
    response = requests.post(url, data=payload)
    access_token = response.json().get("access_token")
    return access_token

# Step 2: Export data using the access token
def export_data(access_token, report_type):
    url = f"https://analyticsapi.zoho.com/api/helmy%40bc-eg.com/Zoho%20Books%20Analytics_1/{report_type}"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Cookie': 'CSRF_TOKEN=250bed86-1a7a-4101-bc59-c7da1b3582f3; JSESSIONID=EC8414F597A7B2134F8FACE58EF59D4D'
    }
    params = {
        'ZOHO_ACTION': 'EXPORT',
        'ZOHO_OUTPUT_FORMAT': 'JSON',
        'ZOHO_ERROR_FORMAT': 'JSON',
        'ZOHO_API_VERSION': '1.0'
    }
    response = requests.get(url, headers=headers, params=params)
    return response.json()

# Step 3: Process the JSON data into a DataFrame
def json_to_dataframe(data):
    try:
        # Extract column headers and rows
        columns = data['response']['result']['column_order']
        rows = data['response']['result']['rows']

        # Convert to DataFrame
        df = pd.DataFrame(rows, columns=columns)

        # Debug: Print columns to check if 'Reporting Date' exists


        # Example of processing the DataFrame: convert 'Reporting Date' to datetime if it exists
        if 'Reporting Date' in df.columns:
            df['Reporting Date'] = pd.to_datetime(df['Reporting Date'], format='%d %b, %Y %H:%M:%S')

        # Additional data processing can be added here as needed
        return df
    except KeyError as e:
        print(f"KeyError: {e} - Check if the column names in the JSON match the expected names.")
        return pd.DataFrame()  # Return an empty DataFrame on error

# Step 4: Save data to Excel
def save_to_excel(df, report_type):
    output = BytesIO()
    df.to_excel(output, index=False, sheet_name=report_type)
    output.seek(0)
    return output

# Function to read and update the sequence number from a file
def get_next_sequence_number(sequence_file):
    if os.path.exists(sequence_file):
        with open(sequence_file, 'r') as f:
            sequence_number = int(f.read().strip())
    else:
        sequence_number = 1  # Start from 1 if the file doesn't exist

    # Update the sequence number for next time
    with open(sequence_file, 'w') as f:
        f.write(str(sequence_number + 1))

    return sequence_number

# Step 5: Send email with multiple attachments
def send_email(file_streams, sequence_number):
    from_email = "yossefsaeed012108@gmail.com"
    to_email = "youssef.shehata@bc-eg.com"

    # Generate the current date in YYYYMMDD format
    current_date = datetime.now().strftime('%Y%m%d')

    # Generate the email subject without the sequence number
    subject = f"[PSI email] Brand Connection_EB_{current_date}"

    body = "Please find the attached data export files."

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    report_types = ["INV", "ST", "SR"]
    for i, file_stream in enumerate(file_streams):
        report_type = report_types[i]
        filename = f"Brand Connection_EB_{report_type}_{current_date}_{str(sequence_number).zfill(2)}.xlsx"
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file_stream.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        msg.attach(part)
        file_stream.seek(0)  # Reset the stream position for the next file

    # Gmail SMTP server configuration
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    password = 'nhxw bozn pxph urio'  # Replace with your email password or app-specific password

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(from_email, password)
        server.send_message(msg)

# Main function
def main():
    sequence_file = "email_sequence.txt"  # File to store the sequence number

    # Uncomment the next line to reset the sequence number to 1
    # reset_sequence_number(sequence_file)

    sequence_number = get_next_sequence_number(sequence_file)  # Get the next sequence number

    access_token = get_access_token()

    # List of report types
    report_types = ["INV", "ST", "SR"]

    file_streams = []

    for report_type in report_types:
        raw_data = export_data(access_token, report_type)
        df = json_to_dataframe(raw_data)  # Convert raw JSON to a processed DataFrame
        file_stream = save_to_excel(df, report_type)
        file_streams.append(file_stream)

    send_email(file_streams, sequence_number)

if __name__ == "__main__":
    main()
