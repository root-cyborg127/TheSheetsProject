import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from PIL import Image, ImageTk
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import re
import webbrowser

class GoogleSheetsCombinerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Google Sheets Combiner")
        self.configure(bg='#2e2e2e')
        self.geometry("600x400")
        self.resizable(False, False)

        self.pages = {}
        self.create_widgets()
        self.show_main_page()

    def create_widgets(self):
        self.create_main_page()
        self.create_input_page()

    def create_main_page(self):
        main_frame = ttk.Frame(self, style='TFrame')
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.pages["main"] = main_frame

        for i in range(5):
            main_frame.grid_rowconfigure(i, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        try:
            self.logo = Image.open("logo.png")
            self.logo = self.logo.resize((100, 100), Image.LANCZOS)
            self.logo = ImageTk.PhotoImage(self.logo)
            logo_label = ttk.Label(main_frame, image=self.logo, style='TLabel')
            logo_label.grid(column=0, row=0, columnspan=2, pady=10)
        except FileNotFoundError:
            messagebox.showerror("Error", "logo.png file not found. Please place it in the script directory.")
            self.quit()

        description_label = ttk.Label(main_frame, text="Google Sheets Combiner", font=("Helvetica", 16), style='TLabel')
        description_label.grid(column=0, row=1, columnspan=2, pady=10)

        num_sheets_label = ttk.Label(main_frame, text="Number of Google Sheets:", style='TLabel')
        num_sheets_label.grid(column=0, row=2, padx=10, pady=10, sticky='e')
        self.num_sheets_entry = ttk.Entry(main_frame, width=5, style="TEntry")
        self.num_sheets_entry.grid(column=1, row=2, padx=10, pady=10, sticky='w')

        next_button = ttk.Button(main_frame, text="Next", command=self.go_to_input_page, style="TButton")
        next_button.grid(column=0, row=3, columnspan=2, pady=10)

        credit_label = ttk.Label(main_frame, text="Developed by Vishwajith Shaijukumar", font=("Helvetica", 8), style='TLabel')
        credit_label.grid(column=0, row=4, columnspan=2, pady=10)

    def create_input_page(self):
        input_frame = ttk.Frame(self, style='TFrame')
        input_frame.grid(row=0, column=0, sticky="nsew")
        self.pages["input"] = input_frame

        label_template = ttk.Label(input_frame, text="Template Sheet ID:", style='TLabel')
        label_template.grid(column=0, row=0, padx=10, pady=10, sticky='e')
        self.entry_template = ttk.Entry(input_frame, width=50, style="TEntry")
        self.entry_template.grid(column=1, row=0, padx=10, pady=10, sticky='w')

        self.sheet_entries = []

        execute_button = ttk.Button(input_frame, text="Execute", command=self.execute_script, style="TButton")
        execute_button.grid(column=0, row=100, columnspan=2, pady=10)

        self.open_link_button = ttk.Button(input_frame, text="Open Link", state=tk.DISABLED, style="TButton")
        self.open_link_button.grid(column=0, row=101, columnspan=2, pady=10)

        self.log_label = ttk.Label(input_frame, text="", wraplength=400, style='TLabel')
        self.log_label.grid(column=0, row=102, columnspan=2, pady=10)

        back_button = ttk.Button(input_frame, text="Back", command=self.show_main_page, style="TButton")
        back_button.grid(column=0, row=103, columnspan=2, pady=10)

        for i in range(104):
            input_frame.grid_rowconfigure(i, weight=1)
        input_frame.grid_columnconfigure(0, weight=1)
        input_frame.grid_columnconfigure(1, weight=1)

    def show_main_page(self):
        self.pages["input"].grid_remove()
        self.pages["main"].grid()

    def show_input_page(self):
        self.pages["main"].grid_remove()
        self.pages["input"].grid()

    def go_to_input_page(self):
        try:
            num_sheets = int(self.num_sheets_entry.get())
            if num_sheets <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid number of sheets.")
            return

        for widget in self.pages["input"].winfo_children():
            widget.grid_remove()

        self.create_input_page()

        for i in range(num_sheets):
            label_url = ttk.Label(self.pages["input"], text=f"Google Sheet URL {i+1}:", style='TLabel')
            label_url.grid(column=0, row=2+i, padx=10, pady=10, sticky='e')
            entry_url = ttk.Entry(self.pages["input"], width=50, style="TEntry")
            entry_url.grid(column=1, row=2+i, padx=10, pady=10, sticky='w')
            self.sheet_entries.append(entry_url)

        self.show_input_page()

    def execute_script(self):
        try:
            self.update_log("Starting execution...")
            SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            SERVICE_ACCOUNT_FILE = 'credentials.json'  # Path to your credentials.json file

            self.update_log("Loading credentials...")
            credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

            self.update_log("Building Google Sheets service...")
            service = build('sheets', 'v4', credentials=credentials)
            drive_service = build('drive', 'v3', credentials=credentials)

            template_sheet_id = self.entry_template.get()

            sheet_urls_and_gids = [(entry.get(), 0) for entry in self.sheet_entries]

            def get_spreadsheet_id(url):
                match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url)
                if match:
                    return match.group(1)
                else:
                    raise ValueError(f"Invalid URL: {url}")

            def get_sheet_data(spreadsheet_id, gid, required_columns):
                sheet = service.spreadsheets()
                result = sheet.get(spreadsheetId=spreadsheet_id).execute()
                sheet_title = [s for s in result['sheets'] if s['properties']['sheetId'] == gid][0]['properties']['title']
                data = sheet.values().get(spreadsheetId=spreadsheet_id, range=sheet_title).execute()
                columns = data['values'][0]
                values = data['values'][1:]

                max_columns = len(columns)
                for row in values:
                    if len(row) < max_columns:
                        row.extend([''] * (max_columns - len(row)))
                    elif len(row) > max_columns:
                        row = row[:max_columns]

                df = pd.DataFrame(values, columns=columns)
                df = df[required_columns]

                return df

            required_columns = ['Due Date', 'Video topic', 'App Promotion', 'Type', 'Thumbnail Text', 'Live Date', 'Status']

            self.update_log("Reading data from all sheets...")
            dataframes = []
            for url, gid in sheet_urls_and_gids:
                try:
                    spreadsheet_id = get_spreadsheet_id(url)
                    df = get_sheet_data(spreadsheet_id, gid, required_columns)
                    dataframes.append(df)
                except Exception as e:
                    self.update_log(f"Error processing sheet: {url} - {e}")

            if not dataframes:
                self.update_log("No valid data found in any sheets.")
                return

            self.update_log("Merging dataframes...")
            combined_df = pd.concat(dataframes, ignore_index=True)
            combined_df = combined_df.fillna('')

            self.update_log("Creating new Google Sheet...")
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

            self.update_log("Uploading data to new sheet...")
            body = {
                'values': [combined_df.columns.values.tolist()] + combined_df.values.tolist()
            }
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range='Sheet1!A1',
                valueInputOption='RAW',
                body=body
            ).execute()

            self.update_log("Copying formatting from template sheet...")
            template_sheet = service.spreadsheets().get(spreadsheetId=template_sheet_id, ranges=['Sheet1'], includeGridData=True).execute()
            template_data = template_sheet['sheets'][0]['data'][0]['rowData']

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

                if requests:
                    service.spreadsheets().batchUpdate(
                        spreadsheetId=spreadsheet_id,
                        body={'requests': requests}
                    ).execute()

            self.update_log("Applying hyperlinks...")
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

            if requests:
                service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={'requests': requests}
                ).execute()

            self.update_log("Sharing the new sheet with anyone who has the link...")
            drive_service.permissions().create(
                fileId=spreadsheet_id,
                body={'type': 'anyone', 'role': 'writer'}  # Make the sheet editable
            ).execute()

            new_sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
            self.update_log(f'Data combined successfully. Access it here: {new_sheet_url}')

            self.open_link_button.config(state=tk.NORMAL, command=lambda: open_link(new_sheet_url))

        except Exception as e:
            self.update_log(f"Error: {str(e)}")

    def update_log(self, message):
        self.log_label.config(text=message)
        self.update_idletasks()

def open_link(url):
    webbrowser.open(url)

if __name__ == "__main__":
    app = GoogleSheetsCombinerApp()

    style = ttk.Style(app)
    style.configure('TFrame', background='#2e2e2e')
    style.configure('TLabel', background='#2e2e2e', foreground='#ffffff', font=('Helvetica', 10))
    style.configure('TEntry', fieldbackground='#333333', foreground='#000000', borderwidth=1, relief='flat')
    style.configure('TButton', background='#444444', foreground='#ffffff', borderwidth=1, relief='flat', padding=6)
    style.map('TButton', background=[('active', '#666666')])

    app.mainloop()
