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

    # Define Tier 1 and Tier 2 sites
    tier_1_sites = [
    "https://www.lib.uchicago.edu/inventory.xml",
    "https://intranet.uchicago.edu/",
    "https://collegeadmissions.uchicago.edu/sitemap.xml",
    "https://safety-security.uchicago.edu/",
    "https://grad.uchicago.edu/",
    "https://giving.uchicago.edu/site/Donation2?df_id=1681&mfc_pref=T&1681.donation=form1",
    "https://news.uchicago.edu/",
    "https://dining.uchicago.edu/",
    "https://www.uchicago.edu/",
    "https://financialaid.uchicago.edu/",
    "https://visit.uchicago.edu/",
    "https://events.uchicago.edu/",
    "https://inclusion.uchicago.edu",
    "https://equalopportunityprograms.uchicago.edu/",
    "https://diversityandinclusion.uchicago.edu/",
    "https://bursar.uchicago.edu/",
    "https://disabilities.uchicago.edu/",
    "https://registrar.uchicago.edu/",
    "https://inclusivepedagogy.uchicago.edu",
    "https://directory.uchicago.edu/",
    "http://portal.uchicago.edu/ais/",
    "https://titleixpolicy.uchicago.edu",
    "https://harassmentpolicy.uchicago.edu/",
    "https://accessibility.uchicago.edu/",
    "https://digitalaccessibility.uchicago.edu/"
    ]

    tier_2_sites = [
    "https://www.classes.cs.uchicago.edu/",
    "https://bfi.uchicago.edu/page-sitemap.xml",
    "https://chas.uchicago.edu/",
    "https://esl.uchicago.edu/",
    "https://athletics.uchicago.edu",
    "https://chemistry.uchicago.edu/",
    "http://medicine.uchicago.edu/",
    "https://graham.uchicago.edu/",
    "https://facilities.uchicago.edu/",
    "https://internationalaffairs.uchicago.edu/",
    "https://divinity.uchicago.edu/",
    "https://global.uchicago.edu/sitemap.xml",
    "https://socialsciences.uchicago.edu/",
    "https://www.ucls.uchicago.edu",
    "https://cehd.uchicago.edu/",
    "https://humanities.uchicago.edu",
    "https://physicalsciences.uchicago.edu/xml-sitemaps/index.xml",
    "https://alumniandfriends.uchicago.edu/s/sfsites/c/resource/sitemaps1/sitemaps1/sitemap-gsc1.xml",
    "https://ccrf.uchicago.edu",
    "https://www.uchicagocharter.org",
    "https://www.journals.uchicago.edu/",
    "https://online.professional.uchicago.edu/",
    "https://music.uchicago.edu/",
    "https://cs.uchicago.edu/",
    "https://www.chicagobooth.edu/executiveeducation/",
    "https://dova.uchicago.edu/",
    "https://arthistory.uchicago.edu/",
    "https://humanresources.uchicago.edu/",
    "https://astrophysics.uchicago.edu/",
    "https://pnf.uchicago.edu/",
    "https://civicengagement.uchicago.edu/",
    "http://giving.uchicago.edu/site/TR?fr_id=1150&pg=entry",
    "https://pme.uchicago.edu/",
    "https://uchicago.hk/",
    "https://classics.uchicago.edu/",
    "https://www.chicagobooth.edu/",
    "https://convocation.uchicago.edu",
    "https://biochem.uchicago.edu",
    "https://cnet.uchicago.edu/CheckIn/CheckInLandingPage",
    "https://aura.uchicago.edu/",
    "https://summer.uchicago.edu",
    "https://ciic.uchicago.edu/",
    "http://rll.uchicago.edu/",
    "https://uchp.uchicago.edu",
    "https://financialeconomics.uchicago.edu/",
    "https://neurosurgery.uchicago.edu",
    "https://physics.uchicago.edu/",
    "https://geosci.uchicago.edu/",
    "http://giving.uchicago.edu/site/TR?sid=1083&type=fr_informational&pg=informational&fr_id=1120",
    "https://press.uchicago.edu/index.html",
    "https://cmes.uchicago.edu/",
    "https://english.uchicago.edu",
    "https://mathematics.uchicago.edu/",
    "https://crownschool.uchicago.edu/",
    "https://professional.uchicago.edu/sitemap.xml",
    "https://courses.uchicago.edu/",
    "https://ura.uchicago.edu/",
    "https://csl.uchicago.edu/",
    "https://safety.uchicago.edu/xml-sitemaps/index.xml",
    "https://slavic.uchicago.edu/",
    "https://familymedicine.uchicago.edu/",
    "https://biosciences.uchicago.edu",
    "https://kavlicosmo.uchicago.edu/",
    "https://mycrownschool.uchicago.edu/",
    "https://cam.uchicago.edu/",
    "https://taps.uchicago.edu/",
    "http://german.uchicago.edu/",
    "https://academictech.uchicago.edu",
    "https://careeradvancement.uchicago.edu",
    "https://cms.uchicago.edu/",
    "https://political-science.uchicago.edu/",
    "https://voices.uchicago.edu/sandbox/",
    "http://nelc.uchicago.edu/",
    "https://finadmin.uchicago.edu/",
    "https://ealc.uchicago.edu/",
    "https://history.uchicago.edu/",
    "https://myphoto.uchicago.edu/",
    "https://linguistics.uchicago.edu/",
    "https://harris.uchicago.edu/",
    "https://www.rockefeller.uchicago.edu/",
    "https://finmath.uchicago.edu/",
    "https://www.law.uchicago.edu/",
    "https://chss.uchicago.edu/",
    "https://humdev.uchicago.edu/",
    "https://salc.uchicago.edu/",
    "https://radiology.uchicago.edu",
    "https://capp.uchicago.edu/",
    "https://sociology.uchicago.edu/",
    "https://www.mbl.edu/",
    "https://economics.uchicago.edu/",
    "https://philosophy.uchicago.edu",
    "https://macss.uchicago.edu/",
    "https://complit.uchicago.edu/",
    "https://cissr.uchicago.edu/",
    "https://stat.uchicago.edu",
    "https://chicagostudies.uchicago.edu",
    "https://maph.uchicago.edu",
    "https://pritzker.uchicago.edu",
    "https://its.uchicago.edu",
    "https://ceas.uchicago.edu/",
    "https://radonc.uchicago.edu",
    "https://anthropology.uchicago.edu/",
    "http://course-info.cs.uchicago.edu/",
    "https://grahamcourses.uchicago.edu/sitemap_gen.xml",
    "https://rdi.uchicago.edu",
    "https://cir.uchicago.edu/",
    "https://startersite.uchicago.edu",
    "https://ggsb.uchicago.edu/",
    "https://pediatrics.uchicago.edu/",
    "https://socialthought.uchicago.edu/",
    "https://psychology.uchicago.edu/",
    "https://help.uchicago.edu",
    "https://researchsafety.uchicago.edu/",
    "https://ophthalmology.uchicago.edu/",
    "https://obgyn.uchicago.edu",
    "https://psychiatry.uchicago.edu/",
    "https://clas.uchicago.edu/",
    "https://surgery.uchicago.edu/",
    "https://mapss.uchicago.edu/",
    "https://politicaleconomy.uchicago.edu/",
    "https://teaching.uchicago.edu/",
    "https://leadforsociety.uchicago.edu/",
    "https://ecologyandevolution.uchicago.edu",
    "https://cnet.uchicago.edu/2FA/index.htm",
    "https://portal.internationalaffairs.uchicago.edu/",
    "https://anesthesia.uchicago.edu/",
    "https://neurology.uchicago.edu/",
    "https://uchicago.cn/",
    "https://mgcb.uchicago.edu/",
    "https://cancerbio.uchicago.edu/",
    "https://immunology.uchicago.edu/",
    "https://publichealth.bsd.uchicago.edu/",
    "https://benmay.uchicago.edu",
    "https://college.uchicago.edu/home",
    "https://oba.bsd.uchicago.edu",
    "https://neurobiology.uchicago.edu",
    "https://microbiology.uchicago.edu",
    "https://ortho.uchicago.edu/",
    "https://africanstudies.uchicago.edu/",
    "https://biologicalsciences.uchicago.edu/",
    "https://drsb.uchicago.edu/",
    "https://genes.uchicago.edu",
    "https://neurograd.uchicago.edu/",
    "https://online.uchicago.edu",
    "https://president.uchicago.edu",
    "https://pathology.uchicago.edu",
    "https://health.uchicago.edu/",
    "https://llso.uchicago.edu",
    "https://provost.uchicago.edu/",
    "https://cns.uchicago.edu/",
    "https://camb.uchicago.edu/",
    "https://chicagoquantum.org/",
    "https://vita.uchicago.edu/",
    "https://uchicago.in/",
    "https://study-abroad.uchicago.edu",
    "https://micro.uchicago.edu/",
    "https://bsdfacilities.uchicago.edu/",
    "https://centerinparis.uchicago.edu",
    "https://ipo.uchicago.edu/",
    "https://medievalstudies.uchicago.edu",
    "https://globalstudies.uchicago.edu",
    "https://pbhs.uchicago.edu/",
    "https://digitalculture.uchicago.edu/teaching/",
    "https://eegraduate.uchicago.edu/",
    "https://evbio.uchicago.edu/",
    "https://hgen.uchicago.edu/",
    "https://metabolism.uchicago.edu/",
    "https://integbio.uchicago.edu/",
    "https://inauguration.uchicago.edu/",
    "https://masterliberalarts.uchicago.edu/",
    "https://disabilitystudies.uchicago.edu/",
    "https://stevanovich.uchicago.edu/",
    "https://arts.uchicago.edu/",
    "https://pediatrics.bsd.uchicago.edu/",
    "https://genome.uchicago.edu/",
    "https://medhum.uchicago.edu/",
    "https://tad.uchicago.edu/",
    "https://aging.uchicago.edu/",
    "https://bmb.uchicago.edu/",
    "https://aamg.uchicago.edu/",
    "https://hphp.uchicago.edu/",
    "https://genomeinformatics.uchicago.edu",
    "https://pme-facilities.uchicago.edu/",
    "https://identity.uchicago.edu/",
    "https://spatialsenses.uchicago.edu/",
    "https://freeexpression.uchicago.edu/",
    "https://thelewisprize.uchicago.edu/",
    "https://artspeaks.uchicago.edu/",
    "https://communication.uchicago.edu",
    "https://myathletics.uchicago.edu/",
    "https://registrar.uchicago.edu/",
    "https://thedog.uchicago.edu/",
    "https://medphys.uchicago.edu/",
    "https://washington.uchicago.edu/",
    "https://cytogenetics.uchicago.edu/",
    "https://datascience.uchicago.edu/",
    "https://energy.uchicago.edu",
    "https://grants.uchicago.edu/",
    "https://planitpurple.uchicago.edu",
    "https://mycourses.uchicago.edu",
    "https://nationalnewschicago.uchicago.edu/",
    "https://newcollegium.uchicago.edu/",
    "https://bitnami-dreamforce8-uc-chicago-test.cloudapp.net/track/en_US/ASTU/ASTU_ThirdYear/blocks/ASTU_Curriculum/fields/ASTU_Curriculum/fields",
    "https://uchicago.bitnamiapp.com/"
    ]

    # Function to categorize tiers
    def categorize_tier(url):
        if url in tier_1_sites:
            return 'Tier 1'
        elif url in tier_2_sites:
            return 'Tier 2'
        else:
            return 'Tier 3'
    
    # Load the parsed CSV file
    parsed_csv = "siteimprove-sheets/parsed_siteimprove_export.csv"
    parsed_df = pd.read_csv(parsed_csv)

    # Categorize the sites into tiers
    parsed_df['Tier'] = parsed_df['URL'].apply(categorize_tier)

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

        # Update the counts (this assumes your formula columns already exist in the sheet)
        sheet.update_cell(2, 4, '=COUNTIF(C:C,"<75")')
        sheet.update_cell(3, 4, '=COUNTIF(C:C,">=75",C:C,"<85")')
        sheet.update_cell(4, 4, '=COUNTIF(C:C,">=85")')

    # Update the Scores by Tier graph (you may need to adjust the range references based on your actual sheet)
    scores_by_tier_sheet = reporting_spreadsheet.worksheet('Scores by Tier')
    scores_by_tier_sheet.update('B2', '=\'Tier 1 scores\'!D4')
    scores_by_tier_sheet.update('C2', '=\'Tier 1 scores\'!D3')
    scores_by_tier_sheet.update('D2', '=\'Tier 1 scores\'!D2')
    scores_by_tier_sheet.update('B3', '=\'Tier 2 scores\'!D4')
    scores_by_tier_sheet.update('C3', '=\'Tier 2 scores\'!D3')
    scores_by_tier_sheet.update('D3', '=\'Tier 2 scores\'!D2')
    scores_by_tier_sheet.update('B4', '=\'Tier 3 scores\'!D4')
    scores_by_tier_sheet.update('C4', '=\'Tier 3 scores\'!D3')
    scores_by_tier_sheet.update('D4', '=\'Tier 3 scores\'!D2')

    print("Issues per month sheet, Progress sheet, Tier sheets, and Scores by Tier graph have been updated.")
except gspread.exceptions.APIError as e:
    print(f"An error occurred with the Google Sheets API: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
