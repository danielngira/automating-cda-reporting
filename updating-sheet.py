import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
from dotenv import load_dotenv
import string
import time

# Load environment variables from .env file
load_dotenv()

# Convert column number to Excel-style column letter
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

    # Open the Google Sheets file for the reporting sheet and access the necessary sheets
    reporting_spreadsheet = client.open_by_key(REPORTING_SHEET_ID)
    issues_per_month_sheet = reporting_spreadsheet.worksheet(ISSUES_PER_MONTH_SHEET)
    progress_sheet = reporting_spreadsheet.worksheet(PROGRESS_SHEET)

    # Convert the "Issues Per Month" sheet to a DataFrame
    issues_data = issues_per_month_sheet.get_all_values()
    issues_columns = issues_data[0]
    issues_df = pd.DataFrame(issues_data[1:], columns=issues_columns)

    # Insert two new columns at positions E and F, shifting existing columns to the right
    issues_per_month_sheet.insert_cols([[], []], col=5)  # Insert at E and F

    # Convert numerical columns to native Python types to avoid JSON serialization issues
    for col in issues_df.columns[1:]:
        try:
            issues_df[col] = pd.to_numeric(issues_df[col])  # Convert to numeric types where applicable
        except (ValueError, TypeError):
            pass  # If conversion fails, leave the column as is
    
    # Add the new month column (e.g., "Issues - Sep24")
    current_month = pd.to_datetime('today').strftime('%b%y')
    new_issues_column = f"Issues - {current_month}"
    remediated_column = f"Issues Remediated - {current_month}"

    # Insert new columns at positions E and F, shifting existing columns to the right
    issues_df.insert(4, new_issues_column, 0)  # Insert the new column at the correct position
    issues_df.insert(5, remediated_column, 0)  # Insert the remediated column, initialize with 0

    # Ensure no NaN values
    issues_df.fillna("", inplace=True)  # Replace NaN with empty strings
    issues_df[new_issues_column] = pd.to_numeric(issues_df[new_issues_column], errors='coerce').fillna(0)

    # Update the "Issues Per Month" sheet with data from the parsed CSV
    parsed_csv = "siteimprove-sheets/parsed_siteimprove_export.csv"
    parsed_df = pd.read_csv(parsed_csv)

    # Validation: Check the number of sites in the CSV
    total_sites_csv = parsed_df.shape[0]
    print(f"Total sites in CSV: {total_sites_csv}")

    # Ensure 'Title' is not NaN for any row in the CSV
    parsed_df = parsed_df.dropna(subset=['Title'])

    # Iterate over parsed CSV and update/add entries in the DataFrame
    for index, row in parsed_df.iterrows():
        site_title = row['Title']
        issues_count = row['Issues']

        # Add or update the site in the DataFrame
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

    # Ensure no missing sites between CSV and the updated sheet
    total_sites_updated = len(issues_df['Title'].unique())  # Count unique Titles after update
    if total_sites_csv != total_sites_updated:
        print(f"Warning: CSV has {total_sites_csv} sites, but only {total_sites_updated} were updated in the sheet.")
    else:
        print(f"Success: All {total_sites_csv} sites have been updated in the sheet.")

    # Write entire DataFrame back to Google Sheets (including any newly added sites)
    issues_per_month_sheet.update([issues_df.columns.values.tolist()] + issues_df.values.tolist(), value_input_option='USER_ENTERED')

    # Prepare updates for columns E and F (new issues and remediated issues)
    updates = []
    for i in range(len(issues_df)):
        row_num = i + 2  # Adjust for the header row
        # Update new issues column (E)
        updates.append({
            'range': f'E{row_num}',
            'values': [[float(issues_df.iloc[i, 4])]]  # Column E is the 5th column in the DataFrame
        })
        # Update remediated issues column (F)
        formula = f"=IF(G{row_num}-E{row_num}<0,0,G{row_num}-E{row_num})"
        updates.append({
            'range': f'F{row_num}',
            'values': [[formula]]
        })

    # Send the batch update to Google Sheets
    issues_per_month_sheet.batch_update(
    [{'range': u['range'], 'values': u['values']} for u in updates],
    value_input_option='USER_ENTERED'
    )

    # Add the total row only to columns E and F without overwriting other columns
    total_row = ['Total', '', '', '']  # Start with the first four empty cells
    for col in issues_df.columns[4:6]:  # Only update columns E and F
        col_letter = colnum_string(issues_df.columns.get_loc(col) + 1)
        total_row.append(f'=SUM({col_letter}2:{col_letter}{len(issues_df) + 1})')

    # Add the total row to columns E and F using batch update
    issues_per_month_sheet.batch_update([
        {'range': f'E{len(issues_df) + 2}', 'values': [[total_row[4]]]},
        {'range': f'F{len(issues_df) + 2}', 'values': [[total_row[5]]]}
    ], value_input_option='USER_ENTERED')

    # Delay to avoid hitting the API quota
    time.sleep(20)  # Increased sleep to prevent API quota issues

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

    # Ensure total_row_number is correctly set after adding the total row
    total_row_number = len(issues_df)

    # Correct formulas in the "Issues remediated to date" and "Issues remaining" columns
    for i in range(2, len(progress_data) + 2):  # Adjust for new row at top
        remediated_col_letter = colnum_string(6 + (i - 2) * 2)  # F starts at column 6
        issues_col_letter = colnum_string(5 + (i - 2) * 2)      # E starts at column 5

        # Adjust the formulas to use the correct row number (total_row_number)
        progress_sheet.update_cell(i, 3, f"=C{i+1}+'Issues per month'!{remediated_col_letter}{total_row_number + 1}")
        progress_sheet.update_cell(i, 4, f"='Issues per month'!{issues_col_letter}{total_row_number + 1}")
        time.sleep(2)  # Increased sleep to avoid API quota issues

    # Extract the first value from 'Tags' to determine the tier
    parsed_df['Tier'] = parsed_df['Tags'].apply(lambda x: x.split(',')[0].strip())

    # Update the Tier sheets and the Scores by Tier sheet
    for tier, sheet_name in TIER_SHEETS.items():
        sheet = reporting_spreadsheet.worksheet(sheet_name)

        # Filter the parsed data for the current tier
        tier_df = parsed_df[parsed_df['Tier'] == tier.split()[-1]]

        # Check if there are any entries for the tier
        if tier_df.empty:
            print(f"No data found for {tier}")
            continue

        # Prepare data for updating
        tier_data = tier_df[['Title', 'URL', 'Accessibility score']].values.tolist()

        # Replace all data in the sheet with new data
        sheet.update([['Name', 'URL', 'Accessibility score']] + tier_data, value_input_option='USER_ENTERED')

    print("Issues per month sheet, Progress sheet, Tier sheets, and Scores by Tier graph have been updated.")

except gspread.exceptions.APIError as e:
    print(f"An error occurred with the Google Sheets API: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")