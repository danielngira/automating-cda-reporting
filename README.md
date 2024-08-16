# SiteImprove Data Automation
===========================

This repository contains a Python script to automate the monthly process of extracting data from SiteImprove CSV files and updating Google Sheets. The automation streamlines the reporting process, reduces manual work, and ensures consistency in data updates.

Table of Contents
-----------------

-   [Features](#features)
-   [Prerequisites](#prerequisites)
-   [Installation](#installation)
-   [Usage](#usage)
-   [Options](#options)
-   [Contributing](#contributing)
-   [License](#license)

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
-   Google Cloud Service Account with access to Google Sheets API
-   Access to the Google Sheets you want to update

### Required Python Libraries

-   `google-auth`
-   `google-auth-oauthlib`
-   `google-auth-httplib2`
-   `google-api-python-client`
-   `pandas`
-   `pyinstaller` (optional for creating an executable)

You can install these libraries using the following command:

bash

Copy code

`pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client pandas pyinstaller`

Installation
------------

1.  **Clone the Repository**:

    bash

    Copy code

    `git clone https://github.com/your-username/siteimprove-automation.git
    cd siteimprove-automation`

2.  **Set Up Google Sheets API**:

    -   Go to the Google Cloud Console.
    -   Enable the **Google Sheets API** for your project.
    -   Create a Service Account and download the JSON credentials file.
    -   Share the Google Sheet with the Service Account email.
3.  **Configure the Script**:

    -   Place your Google Sheets API credentials JSON file in the project directory.
    -   Update the script with your Google Sheet ID and any specific ranges you wish to modify.

Usage
-----

### Running the Script

You can run the script directly with Python:

bash

Copy code

`python siteimprove_automation.py`

### Creating an Executable

To create an executable, use PyInstaller:

bash

Copy code

`pyinstaller --onefile siteimprove_automation.py`

### Scheduling the Script

You can schedule the script to run automatically:

-   **Windows**: Use Task Scheduler.
-   **macOS/Linux**: Use a cron job.

### Web Interface (Optional)

You can also set up a simple web interface using Flask to allow for file uploads and script execution.

Options
-------

-   **Executable File**: Convert the script into an executable for easy distribution.
-   **Scheduled Task**: Schedule the script to run at specific intervals.
-   **Web Interface**: Set up a Flask-based web interface for file uploads and updates.

Contributing
------------

Contributions are welcome! Please fork this repository and submit a pull request with your changes.

License
-------

This project is licensed under the MIT License. See the `LICENSE` file for more details.

* * * * *

Replace `"your-username"` in the repository URL with your actual GitHub username. This README should give users a comprehensive overview of your project and instructions on how to set it up and use it. Let me know if you need any additional customization!
