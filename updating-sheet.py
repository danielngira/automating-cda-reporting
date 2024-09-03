import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
from dotenv import load_dotenv
import string
import time

load_dotenv()

def colnum_string(n):
    """Convert a column number to an Excel-style column letter."""
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

try:
    # Google Sheets setup
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SERVICE_ACCOUNT_FILE = 'siteimprove-sheets/service-account.json'
    credentials = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, SCOPES)
    client = gspread.authorize(credentials)

    # Google Sheets IDs and Range
    REPORTING_SHEET_ID = os.getenv("REPORTING_SHEET_ID")

    ISSUES_PER_MONTH_SHEET = 'Issues per month'
    PROGRESS_SHEET = 'Progress'
    TIER_SHEETS = {
        'Tier 1': 'Tier 1 scores',
        'Tier 2': 'Tier 2 scores',
        'Tier 3': 'Tier 3 scores'
    }

    # Function to categorize tiers based on the first tag
    def categorize_tier(tags):
        tier = tags.split(',')[0].strip()  # Extract the first tag and remove any leading/trailing spaces
        if tier == '1':
            return 'Tier 1'
        elif tier == '2':
            return 'Tier 2'
        else:
            return 'Tier 3'
    
    # Load the parsed CSV file
    parsed_csv = "siteimprove-sheets/parsed_siteimprove_export.csv"
    parsed_df = pd.read_csv(parsed_csv)

    # Categorize the sites into tiers
    parsed_df['Tier'] = parsed_df['Tags'].apply(categorize_tier)

    # Open the Google Sheets file for the reporting sheet and access the necessary sheets
    reporting_spreadsheet = client.open_by_key(REPORTING_SHEET_ID)
    issues_per_month_sheet = reporting_spreadsheet.worksheet(ISSUES_PER_MONTH_SHEET)
    progress_sheet = reporting_spreadsheet.worksheet(PROGRESS_SHEET)

    # Convert the "Issues Per Month" sheet to a DataFrame
    issues_data = issues_per_month_sheet.get_all_values()
    issues_columns = issues_data[0]
    issues_df = pd.DataFrame(issues_data[1:], columns=issues_columns)

    # Add the new month column (e.g., "Issues - Aug24")
    current_month = pd.to_datetime('today').strftime('%b%y')
    new_issues_column = f"Issues - {current_month}"
    issues_df.insert(4, new_issues_column, 0)  # Insert the new column at the correct position

    # Calculate the previous month string and column name
    previous_month = pd.to_datetime('today') - pd.DateOffset(months=1)
    previous_month_str = previous_month.strftime('%b%y')
    previous_issues_column = f"Issues - {previous_month_str}"

    # Add the remediated column (e.g., "Issues Remediated - Jul24")
    remediated_column = f"Issues Remediated - {previous_month_str}"
    issues_df.insert(5, remediated_column, '')  # Insert the remediated column, initially empty

    # Update the "Issues Per Month" sheet with data from the parsed CSV
    for index, row in parsed_df.iterrows():
        site_title = row['Title']
        issues_count = row['Issues']

        if site_title in issues_df['Title'].values:
            # Update existing site
            issues_df.loc[issues_df['Title'] == site_title, new_issues_column] = issues_count
        else:
            # Add new site by concatenating a new DataFrame
            new_row = pd.DataFrame({
                'Title': [site_title],
                'URL': [row['URL']],
                'Tags': [row['Tags']],
                new_issues_column: [issues_count],
                remediated_column: [0]  # Default to 0 for new entries
            })
            issues_df = pd.concat([issues_df, new_row], ignore_index=True)

    # Ensure columns are numeric before performing calculations
    issues_df[new_issues_column] = pd.to_numeric(issues_df[new_issues_column], errors='coerce').fillna(0)
    if previous_issues_column in issues_df.columns:
        issues_df[previous_issues_column] = pd.to_numeric(issues_df[previous_issues_column], errors='coerce').fillna(0)

    # Calculate the remediated issues for the previous month
    if previous_issues_column in issues_df.columns:
        for i in range(len(issues_df)):
            prev_col_letter = colnum_string(issues_df.columns.get_loc(previous_issues_column) + 1)
            new_col_letter = colnum_string(issues_df.columns.get_loc(new_issues_column) + 1)
            row_num = i + 2  # Adjust for the header row and 0-based index
            formula = (
                f"=IF({prev_col_letter}{row_num}-{new_col_letter}{row_num}<0,"
                f"0,{prev_col_letter}{row_num}-{new_col_letter}{row_num})"
            )
            issues_df.iloc[i, issues_df.columns.get_loc(remediated_column)] = formula

    # Remove sites that are no longer in the parsed CSV
    current_sites = parsed_df['Title'].tolist()
    issues_df = issues_df[issues_df['Title'].isin(current_sites)]

    # Sort the DataFrame alphabetically by Title
    issues_df = issues_df.sort_values(by='Title')

    # Replace NaN with 0 or an appropriate value
    issues_df = issues_df.fillna(0)

    # Add a total row
    total_row = ['Total', '', '', '']  # Start with the first four empty cells
    for col in issues_df.columns[4:]:
        col_letter = colnum_string(issues_df.columns.get_loc(col) + 1)
        total_row.append(f'=SUM({col_letter}2:{col_letter}{len(issues_df) + 1})')  # Create a formula for each column

    issues_df.loc[len(issues_df)] = total_row

    # Write the updated DataFrame back to the "Issues Per Month" Google Sheet
    issues_per_month_sheet.update([issues_df.columns.values.tolist()] + issues_df.values.tolist(), value_input_option='USER_ENTERED')

    # Delay to avoid hitting the API quota
    time.sleep(20)  # Increased sleep to prevent API quota issues

    # Refresh the "Issues Per Month" data after update
    issues_data = issues_per_month_sheet.get_all_values()
    issues_columns = issues_data[0]
    issues_df = pd.DataFrame(issues_data[1:], columns=issues_columns)

    # Calculate the total row number again after updating
    total_row_number = len(issues_df) + 1

    # Now update the Progress sheet after updating Issues per month sheet
    progress_data = progress_sheet.get_all_values()
    next_month = (pd.to_datetime(progress_data[1][0]) + pd.DateOffset(months=1)).strftime('%m/1/%y')

    # Calculate Sites Reviewed
    total_sites_reviewed = len(issues_df) - 1  # Exclude the Total row itself

    # Insert a new row at the top of the "Progress" sheet
    new_row = [
        next_month,  # New date
        total_sites_reviewed,  # Sites reviewed
        "",  # Issues remediated to date (will be updated below)
        ""   # Issues remaining (will be updated below)
    ]

    # Insert the new row at the top
    progress_sheet.insert_row(new_row, index=2)

    # Correct formulas in the "Issues remediated to date" and "Issues remaining" columns
    for i in range(2, len(progress_data) + 2):  # Adjust for new row at top
        remediated_col_letter = colnum_string(6 + (i - 2) * 2)  # F starts at column 6
        issues_col_letter = colnum_string(5 + (i - 2) * 2)      # E starts at column 5

        # Adjust the formulas to use the correct row number (total_row_number)
        progress_sheet.update_cell(i, 3, f"=C{i+1}+'Issues per month'!{remediated_col_letter}{total_row_number}")
        progress_sheet.update_cell(i, 4, f"='Issues per month'!{issues_col_letter}{total_row_number}")
        time.sleep(2)  # Increased sleep to avoid API quota issues

    # The logic to update the Tier sheets and the Scores by Tier sheet

    # Update each tier sheet with the corresponding data
    for tier, sheet_name in TIER_SHEETS.items():
        sheet = reporting_spreadsheet.worksheet(sheet_name)

        # Filter the parsed data for the current tier
        tier_df = parsed_df[parsed_df['Tier'] == tier]

        # Prepare data for updating
        tier_data = tier_df[['Title', 'URL', 'Accessibility score']].values.tolist()

        # Replace all data in the sheet with new data
        sheet.update([['Name', 'URL', 'Accessibility score']] + tier_data, value_input_option='USER_ENTERED')

    print("Issues per month sheet, Progress sheet, Tier sheets, and Scores by Tier graph have been updated.")
except gspread.exceptions.APIError as e:
    print(f"An error occurred with the Google Sheets API: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
