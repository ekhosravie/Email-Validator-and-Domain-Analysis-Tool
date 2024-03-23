import pandas as pd
import requests
import json
import logging
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Load Excel file
excel_file = r'D:\\project\\100 Emails.xlsx'

# Read the Excel file into a DataFrame
df = pd.read_excel(excel_file)

# Read the 'Keywords for matching' sheet from the Excel file
df_keyword = pd.read_excel(excel_file, sheet_name='Keywords for matching' ,header=None)

# Print the DataFrame containing keywords
print(df_keyword)

# API key for Mailboxlayer API
api_key = '8f855b8abc15231388d370521'

# Sample email address for testing
email = 'e.khosravi@gmail.com'

# Function to validate emails via the Mailboxlayer API
def validate_email(email):
    api_key = 'your_mailboxlayer_api_key'  # Replace with your Mailboxlayer API key
    url = f'https://apilayer.net/api/check?access_key={api_key}&email={email}'
    response = requests.get(url)
    data = response.json()
    return data

# Function to extract email domains and perform keyword matching
def analyze_email(email, keywords):
    validation_result = validate_email(email)
    domain = email.split('@')[-1]
    matched = False
    for keyword in keywords:
        if keyword.lower() in domain.lower():
            matched = True
            break
    return {'email': email, 'domain': domain, 'matched': matched}

# Load keywords for matching
keywords = df_keyword['Keywords for matching'].tolist()

# Validate and analyze emails
validated_emails = []
for index, row in df['All emails'].iteritems():
    analysis_result = analyze_email(row, keywords)
    validated_emails.append(analysis_result)

# Path to the service account key JSON file
service_account_key_json = 'D:\\project\\emailvalidation-417909-e71c84823113.json'

# Function to upload analyzed data to Google Sheets
def upload_to_google_sheets(data):
    try:
        # Set up Google Sheets API client
        credentials = service_account.Credentials.from_service_account_file(service_account_key_json)
        service = build('sheets', 'v4', credentials=credentials)

        # Define spreadsheet ID and ranges for each sheet
        spreadsheet_id = '1vMaXt6sCSznYjsm_GQpCQVX1edULcqjasSAA32qIFQ0'
        sheet_ranges = {
            'email addresses': 'Sheet1!A:B',
            'Valid Email Addresses': 'Valid Email Addresses!A:A',
            'Invalid Email Addresses': 'Invalid Email Addresses!A:A'
        }

        # Initialize data for each sheet
        email_addresses = []
        valid_emails = []
        invalid_emails = []

        # Separate emails into valid and invalid lists
        for d in data:
            email_addresses.append([d['email'], 'Valid' if d['matched'] else 'Invalid'])
            if d['matched']:
                valid_emails.append([d['email']])
            else:
                invalid_emails.append([d['email']])

        # Prepare data for each sheet
        data_for_sheets = {
            'email addresses': email_addresses,
            'Valid Email Addresses': valid_emails,
            'Invalid Email Addresses': invalid_emails
        }

        # Update each sheet with its respective data
        for sheet_name, range_name in sheet_ranges.items():
            body = {'values': data_for_sheets[sheet_name]}
            result = service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id, range=range_name, valueInputOption='RAW', body=body).execute()
            logger.info(f'Data uploaded to Google Sheets ({sheet_name}): {result}')

    except HttpError as error:
        logger.error(f'An error occurred: {error}')

# Upload validated emails to Google Sheets
upload_to_google_sheets(validated_emails)
