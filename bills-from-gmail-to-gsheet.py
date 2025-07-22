from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import pickle
import os
from bs4 import BeautifulSoup
import base64
import re
from email.utils import parsedate_to_datetime
import pyperclip
from dotenv import load_dotenv

load_dotenv()

SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/spreadsheets'
]

SHEET_ID = os.getenv('SHEET_ID')
if not SHEET_ID:
    raise ValueError("SHEET_ID not found in environment variables")
GMAIL_QUERY = os.getenv('GMAIL_QUERY')
if not GMAIL_QUERY:
    raise ValueError("GMAIL_QUERY not found in environment variables")
SHEET_RANGE = os.getenv('SHEET_RANGE', )
if not SHEET_RANGE:
    raise ValueError("SHEET_RANGE not found in environment variables")


def get_services():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file('bills-from-email-to-gsheet-credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    gmail_service = build('gmail', 'v1', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)

    return gmail_service, sheets_service

def extract_html_payload(payload):
    """Recursively get the HTML part from the email payload."""
    if payload.get('mimeType') == 'text/html':
        data = payload['body'].get('data')
        if data:
            return base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')

    elif payload['mimeType'].startswith('multipart'):
        for part in payload.get('parts', []):
            html = extract_html_payload(part)
            if html:
                return html
    return None

def find_order_total_from_html(html):
    soup = BeautifulSoup(html, 'html.parser')
    
    # Find the p tag that contains "ORDER TOTAL:"
    order_total_label = soup.find("p", text=re.compile("ORDER TOTAL", re.I))
    if not order_total_label:
        return None

    # Step 2: Navigate to the parent <tr>, then to the next <tr>
    tr = order_total_label.find_parent('tr')
    next_tr = tr.find_next_sibling('tr') if tr else None
    if not next_tr:
        return None

    # Step 3: Find the <p> inside next <tr>
    price_p = next_tr.find('p')
    if not price_p:
        return None

    # Step 4: Clean & return text
    price = price_p.get_text(strip=True).replace('\u200c', '')
 
    return price


def parse_enercare_receipt(html):
    soup = BeautifulSoup(html, 'html.parser')
    
    # Step 1: Locate the correct table that contains the receipt
    target_table = None
    for table in soup.find_all("table"):
        if table.find(string=lambda s: s and "Your payment receipt:" in s):
            target_table = table
            break

    if not target_table:
        print("❌ Receipt table not found.")
        return {}

    # Step 2: Define the fields we want
    receipt_headers = [
        "ORDER DATE",
        "BILLING ACCOUNT NUMBER",
        "PAYMENT REFERENCE ID",
        "ORDER TOTAL",
        "PAYMENT METHOD"
    ]

    # Step 3: Walk through <tr> tags
    trs = target_table.find_all("tr")
    data = {}

    i = 0
    while i < len(trs) - 1:
        label_td = trs[i].find("td")
        if label_td:
            label_text = label_td.get_text(strip=True).rstrip(":").upper()
            if label_text in receipt_headers:
                value_td = trs[i + 1].find("td")
                if value_td:
                    data[label_text] = value_td.get_text(strip=True).replace('\u200c', '')
                i += 1  # skip next row since it's the value
        i += 1

    print(data)
    return data



def fetch_emails(gmail_service, query='label:inbox', max_results=5):
    results = gmail_service.users().messages().list(userId='me', q=query, maxResults=max_results).execute()
    messages = results.get('messages', [])

    emails = []
    for msg in messages:
        msg_data = gmail_service.users().messages().get(userId='me', id=msg['id']).execute()
        headers = msg_data['payload']['headers']
        subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '')
        date = next((h['value'] for h in headers if h['name'] == 'Date'), '')

        # Get the body content (HTML or plain text)
        parts = msg_data['payload'].get('parts', [])
        for part in parts:
            if part['mimeType'] == 'text/html':
                body_data = part['body'].get('data')
                
                decoded = base64.urlsafe_b64decode(body_data).decode('utf-8', errors='ignore')
                order_total = find_order_total_from_html(decoded)
                
                parse_enercare_receipt(decoded)
                break

        # Parse into datetime object
        dt = parsedate_to_datetime(date)
        # Format it to "21 Mar 2025"
        formatted_date = dt.strftime('%d %b %Y')
        item_name = 'enercare'


        emails.append([item_name, formatted_date, order_total, subject])


    return emails

def write_to_sheet(sheets_service, sheet_id, emails, sheet_range=SHEET_RANGE):
    
    
    """
    Appends new email records to a Google Sheet and sorts the sheet by date.

    Args:
        sheets_service (Any): An authenticated Google Sheets API service instance.
        sheet_id (str): The ID of the target Google Sheet.
        emails (List[List[str]]): A list of email data rows to write, each containing:
            [item_name, formatted_date, order_total, subject].
        sheet_range (str, optional): The A1 range to read from and append to. Defaults to 'Sheet1!A1'.

    Returns:
        None
    """
    
    sheet = sheets_service.spreadsheets()

    
    
    # Step 1: Get existing rows to prevent duplicates
    result = sheet.values().get(spreadsheetId=sheet_id, range=sheet_range).execute()
    
    existing_rows = result.get('values', [])[1:]  # Skip header row

    # Step 2: Build a set of existing records (joined string for uniqueness)
    existing_set = set(tuple(row) for row in existing_rows)

    # Step 3: Filter only new rows
    new_rows = [row for row in emails if tuple(row) not in existing_set]

    if not new_rows:
        print("✅ No new entries to append.")
        return

    # Step 4: Append new rows
    append_range = f"{sheet_range}"
    sheet.values().append(
        spreadsheetId=sheet_id,
        range=append_range,
        valueInputOption='USER_ENTERED',
        body={'values': new_rows}
    ).execute()
    print(f"✅ Appended {len(new_rows)} new rows.")

    # Step 5: Sort entire sheet by date (assumes "date" is column B = index 1)
    sort_request = {
        "requests": [{
            "sortRange": {
                "range": {
                    "sheetId": 0,  # Assuming your data is in the first sheet
                    "startRowIndex": 1,  # Skip header
                    "startColumnIndex": 0,
                    "endColumnIndex": 4
                },
                "sortSpecs": [{
                    "dimensionIndex": 1,  # Column B = date
                    "sortOrder": "ASCENDING"
                }]
            }
        }]
    }

    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body=sort_request
    ).execute()


def main():
    gmail_service, sheets_service = get_services()
    enercare_receipt_emails = fetch_emails(gmail_service, query=GMAIL_QUERY)
    # write_to_sheet(sheets_service, sheet_id=SHEET_ID, emails=enercare_receipt_emails)

if __name__ == '__main__':
    main()

