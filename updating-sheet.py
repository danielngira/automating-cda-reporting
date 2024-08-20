import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
from dotenv import load_dotenv, dotenv_values
load_dotenv()

try:
    # Google Sheets setup
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SERVICE_ACCOUNT_FILE = 'siteimprove-sheets/service-account.json'
    credentials = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, SCOPES)
    client = gspread.authorize(credentials)

    # Google Sheets IDs and Range
    REPORTING_SHEET_ID = os.getenv("REPORTING_SHEET_ID")
    INITIAL_CRAWL_SHEET_ID = os.getenv("INITIAL_CRAWL_SHEET_ID")

    ISSUES_PER_MONTH_SHEET = 'Issues per month'
    INITIAL_CRAWL_SHEET = 'Initial scores'

    # Load the parsed CSV file
    parsed_csv = "siteimprove-sheets/parsed_siteimprove_export.csv"
    parsed_df = pd.read_csv(parsed_csv)

    # Open the Google Sheets file for the reporting sheet and access the necessary sheet
    reporting_spreadsheet = client.open_by_key(REPORTING_SHEET_ID)
    issues_per_month_sheet = reporting_spreadsheet.worksheet(ISSUES_PER_MONTH_SHEET)

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
    issues_df.insert(5, remediated_column, 0)  # Insert the remediated column

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
        issues_df[remediated_column] = issues_df[previous_issues_column] - issues_df[new_issues_column]
        # Ensure no negative values
        issues_df[remediated_column] = issues_df[remediated_column].apply(lambda x: max(x, 0))
    else:
        issues_df[remediated_column] = 0  # Default to 0 if no previous issues column exists

    # Remove sites that are no longer in the parsed CSV
    current_sites = parsed_df['Title'].tolist()
    issues_df = issues_df[issues_df['Title'].isin(current_sites)]

    # Replace NaN with 0 or an appropriate value
    issues_df = issues_df.fillna(0)

    # Write the updated DataFrame back to the "Issues Per Month" Google Sheet
    issues_per_month_sheet.update([issues_df.columns.values.tolist()] + issues_df.values.tolist())

    # Open the Google Sheets file for the "Initial Crawl" sheet
    initial_crawl_spreadsheet = client.open_by_key(INITIAL_CRAWL_SHEET_ID)
    initial_crawl_sheet = initial_crawl_spreadsheet.worksheet(INITIAL_CRAWL_SHEET)

    # Convert the "Initial Crawl" sheet to a DataFrame
    initial_crawl_data = initial_crawl_sheet.get_all_values()
    initial_crawl_columns = initial_crawl_data[0]
    initial_crawl_df = pd.DataFrame(initial_crawl_data[1:], columns=initial_crawl_columns)

    # Remove sites that are no longer in the parsed CSV from the "Initial Crawl" sheet
    initial_crawl_df = initial_crawl_df[initial_crawl_df['Site'].isin(current_sites)]

    # Replace NaN with 0 or an appropriate value
    initial_crawl_df = initial_crawl_df.fillna(0)

    # Write the updated DataFrame back to the "Initial Crawl" Google Sheet
    initial_crawl_sheet.update([initial_crawl_df.columns.values.tolist()] + initial_crawl_df.values.tolist())

    print("Reporting sheet and Initial Crawl sheet have been updated.")
except gspread.exceptions.APIError as e:
    print(f"An error occurred with the Google Sheets API: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
