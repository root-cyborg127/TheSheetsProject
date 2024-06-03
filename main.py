import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import re

# Define the scope and authenticate with Google Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'credentials.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Create the Sheets API service
service = build('sheets', 'v4', credentials=credentials)
drive_service = build('drive', 'v3', credentials=credentials)

# Template Google Sheet ID
template_sheet_id = '1vKGjs9krZfO4_iM0Myz5pAduFDX_7UnPlSxw-pL7-jo'

# Define the URLs and GIDs of the sheets to combine
sheet_urls_and_gids = [
    ("https://docs.google.com/spreadsheets/d/1yZAQxDyzAtcpzjwz89Y97VTLjcnADc6b803RfxEMF-w/edit#gid=0", 0),
    ("https://docs.google.com/spreadsheets/d/1ve61u_Z46H3Pvhyr_UwLKsv2MqKpNu2f9M81p2yj3E8/edit#gid=0", 0),
    ("https://docs.google.com/spreadsheets/d/1554WbYJCaenrKqDTZHLmfLq9SZrjtxWYpqyeGGjmpW4/edit#gid=0", 0),
    ("https://docs.google.com/spreadsheets/d/1cKjHVtQVOTz89zYypysUj5rV9ZdhPbfT15rv6tUZ62Y/edit#gid=0", 0)
]

# Helper function to extract spreadsheet ID from URL
def get_spreadsheet_id(url):
    return re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url).group(1)

# Helper function to get sheet data as a DataFrame
def get_sheet_data(spreadsheet_id, gid, required_columns):
    sheet = service.spreadsheets()
    result = sheet.get(spreadsheetId=spreadsheet_id).execute()
    sheet_title = [s for s in result['sheets'] if s['properties']['sheetId'] == gid][0]['properties']['title']
    data = sheet.values().get(spreadsheetId=spreadsheet_id, range=sheet_title).execute()
    columns = data['values'][0]
    values = data['values'][1:]
    
    # Ensure each row has the same number of columns as the header
    max_columns = len(columns)
    for row in values:
        if len(row) < max_columns:
            row.extend([''] * (max_columns - len(row)))
        elif len(row) > max_columns:
            row = row[:max_columns]
    
    df = pd.DataFrame(values, columns=columns)
    
    # Select only required columns
    df = df[required_columns]
    
    return df

# Required columns
required_columns = ['Due Date', 'Video topic', 'App Promotion', 'Type', 'Thumbnail Text', 'Live Date', 'Status']

# Read data from all sheets
dataframes = []
for url, gid in sheet_urls_and_gids:
    spreadsheet_id = get_spreadsheet_id(url)
    df = get_sheet_data(spreadsheet_id, gid, required_columns)
    dataframes.append(df)

# Merge dataframes
combined_df = pd.concat(dataframes, ignore_index=True)

# Replace NaN values with empty strings
combined_df = combined_df.fillna('')

# Create a new Google Sheet
spreadsheet_body = {
    'properties': {
        'title': 'Combined Sheet Based on Template'
    },
    'sheets': [
        {
            'properties': {
                'title': 'Sheet1'
            }
        }
    ]
}
spreadsheet = service.spreadsheets().create(body=spreadsheet_body, fields='spreadsheetId,sheets').execute()
spreadsheet_id = spreadsheet.get('spreadsheetId')
new_sheet_id = spreadsheet['sheets'][0]['properties']['sheetId']

# Prepare data for upload
body = {
    'values': [combined_df.columns.values.tolist()] + combined_df.values.tolist()
}
service.spreadsheets().values().update(
    spreadsheetId=spreadsheet_id,
    range='Sheet1!A1',
    valueInputOption='RAW',
    body=body
).execute()

# Copy formatting from the template sheet
template_sheet = service.spreadsheets().get(spreadsheetId=template_sheet_id, ranges=['Sheet1'], includeGridData=True).execute()
template_data = template_sheet['sheets'][0]['data'][0]['rowData']

# Check if we have any rows to format
if template_data:
    requests = []
    for i, row in enumerate(template_data):
        if 'values' in row:
            for j, cell in enumerate(row['values']):
                requests.append({
                    'updateCells': {
                        'rows': [{
                            'values': [cell]
                        }],
                        'fields': 'userEnteredFormat.backgroundColor,userEnteredFormat.textFormat',
                        'range': {
                            'sheetId': new_sheet_id,
                            'startRowIndex': i,
                            'endRowIndex': i + 1,
                            'startColumnIndex': j,
                            'endColumnIndex': j + 1
                        }
                    }
                })

    # Execute the batch update if there are formatting requests
    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': requests}
        ).execute()

# Apply hyperlinks
requests = []
for i, row in combined_df.iterrows():
    for j, col in enumerate(required_columns):
        if 'http' in str(row[col]):
            requests.append({
                'updateCells': {
                    'rows': [{
                        'values': [{
                            'userEnteredValue': {
                                'stringValue': row[col]
                            },
                            'textFormatRuns': [{
                                'startIndex': 0,
                                'format': {
                                    'link': {
                                        'uri': row[col]
                                    }
                                }
                            }]
                        }]
                    }],
                    'fields': 'userEnteredValue,textFormatRuns',
                    'range': {
                        'sheetId': new_sheet_id,
                        'startRowIndex': i + 1,
                        'endRowIndex': i + 2,
                        'startColumnIndex': j,
                        'endColumnIndex': j + 1
                    }
                }
            })

# Execute the batch update if there are hyperlink requests
if requests:
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': requests}
    ).execute()

# Share the new sheet with anyone who has the link
drive_service.permissions().create(
    fileId=spreadsheet_id,
    body={'type': 'anyone', 'role': 'reader'}
).execute()

# Print the URL of the new sheet
new_sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
print(f'Data combined successfully into "Combined Sheet". You can access it here: {new_sheet_url}')

