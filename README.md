# SiteImprove Data Automation

Overview
--------

This project aims to automate the process of updating Google Sheets with data extracted from Siteimprove reports. The automation script parses the Siteimprove CSV export, updates specific sheets with the relevant data, and ensures that the data is organized and formatted correctly for reporting purposes.


Table of Contents
-----------------

  [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
  - [Required Python Libraries](#required-python-libraries)
- [Service Account Setup](#service-account-setup)
- [Scripts](#scripts)
  - [`siteimprove-parser.py`](#siteimprove-parserpy)
  - [`updating-sheet.py`](#updating-sheetpy)
- [Environment Variables](#environment-variables)
- [Troubleshooting](#troubleshooting)
- [Next Steps](#next-steps)
- [Running the code](#running-the-code)

Features
--------

-   **CSV Parsing**: Automatically loads and processes the CSV export from SiteImprove.
-   **Google Sheets Update**: Updates the "Issues Per Month" tab in the Google Sheet with new data.
-   **Progress Tracking**: Automatically updates the progress graph in the Google Sheet.
-   **User-Friendly Execution**: Options to run the script as an executable, schedule it with a task, or use a web interface.

Prerequisites
-------------

Before running the script, ensure you have the following installed:

-   Python 3.x
-   `pip` (Python package installer)

### Required Python Libraries

- `google_api_python_client`
- `pandas`
- `protobuf`
- `google-api-core`
- `google-api-python-client`
- `google-auth`
- `google-auth-httplib2`
- `google-auth-oauthlib`
- `googleapis-common-protos`
- `httplib2`
- `gspread`
- `oauth2client`
- `python-dateutil`
- `python-dotenv`
- `pyinstaller` (optional for creating an executable)

You can install these libraries using the following command:

`pip install -r /path/to/requirements.txt`

Service Account Setup
---------------------

To use the Google Sheets API, you must set up a Google Cloud service account:

1. Go to the Google Cloud Console and create a new project.
2. Navigate to "APIs & Services" > "Credentials" and create a new service account.
3. Download the `service-account.json` file and place it in the `siteimprove-sheets` directory.

Scripts
-------

### `siteimprove-parser.py`

This script processes the raw Siteimprove CSV export by:

- Selecting only the necessary columns: `Title`, `URL`, `Tags`, `Date added`, `Accessibility score`, and `Issues`.
- Converting the 'Date added' column to a valid datetime format.
- Sorting the data alphabetically by `Title`.
- Saving the cleaned and sorted data to `parsed_siteimprove_export.csv`.

### `updating-sheet.py`

This script automates the update process for Google Sheets by:

- Inserting new data into the "Issues per month" sheet with the current month's data.
- Calculating the remediated issues by comparing the current and previous month's data.
- Adding a total row to the sheet.
- Updating the "Progress" sheet with the latest sites reviewed, issues remediated to date, and issues remaining.
- Ensuring formulas in the Progress sheet correctly reference the updated totals.

Environment Variables
---------------------

Ensure that the `.env` file is correctly set up with the following variables:

```
REPORTING_SHEET_ID="your-reporting-sheet-id"
INITIAL_CRAWL_SHEET_ID="your-initial-crawl-sheet-id"
```

Troubleshooting
---------------

- **Issue**: "API quota exceeded" error.
  - **Solution**: Increase the delay between API calls in the script, or reduce the frequency of script execution.
- **Issue**: "Module not found" error when running the script.
  - **Solution**: Ensure all dependencies are installed by running `pip install -r requirements.txt`.

Next Steps
----------

1. **Enhancements:**

    - Generate a unique name for the parsed file each time it is created.
    - Add the parsed file to `.gitignore` to prevent it from being committed.
2. **Service Account:**

    - Transition from using personal Google Cloud credentials to CDA's Google Cloud credentials.
3. **Dependencies:**

    - Review and update `requirements.txt` to ensure all necessary Python packages are included.
4. **Sheet IDs:**

    - Replace the test Google Sheets IDs with the actual IDs.
    - Share the updated sheet IDs with Jack and Thalia.
5. **Reports & Graphs:**

    - Update the graph and its headers within the reporting sheet.
    - Update the Scores by tier in the relevant sheet.
6. **Packaging:**

    - Package the entire application for easy deployment and sharing with the team.

Running the code
----------

- I am working with test sheets in `Student worker created` -> `Siteimprove reporting test sheet` and `Test Initial Crawl a11y scores`
- In the `Issues per month` sheet, delete column E and F
- In the `Progress` sheet, delete row 2

- Run:

```
python path/to/siteimprove-parser.py
python3 path/to/updating-sheet.py
```

- If these are successful, you will get two messages, one saying "Parsed and sorted CSV has been saved" and another saying "Issues per month sheet and Progress sheet have been updated."

- The script to update the sheets may take some time because I added a delay to avoid hitting API quota limits.
