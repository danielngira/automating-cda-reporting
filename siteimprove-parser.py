import pandas as pd
from dateutil import parser

# Load the Siteimprove CSV, specifying the correct columns to load
input_csv = "siteimprove-sheets/siteimprove_export.csv"
col_names = ['Title', 'URL', 'Tags', 'Date added', 'Accessibility score', 'Site target', 'Points to target', 'Pages', 'Issues', 'Potential issues', 'PDFs with issues']
df = pd.read_csv(input_csv, encoding='utf-16', skiprows=3, delimiter='\t', names=col_names)

# Select only the necessary columns
df = df[['Title', 'URL', 'Tags', 'Date added', 'Accessibility score', 'Issues']]

# Convert 'Date added' to string and strip any leading/trailing whitespace
df['Date added'] = df['Date added'].astype(str).str.strip()

# Manually parse the dates using dateutil.parser.parse
def parse_date(date_str):
    try:
        return parser.parse(date_str)
    except parser.ParserError:
        return pd.NaT

df['Date added'] = df['Date added'].apply(parse_date)

# Identify rows where the conversion failed
invalid_dates = df[df['Date added'].isna()]
print("Rows with invalid date format after parsing:")
print(invalid_dates)

# Drop rows where date conversion failed
df = df.dropna(subset=['Date added'])

# Filter based on the last month's date range
last_month = pd.to_datetime('today').normalize() - pd.DateOffset(months=1)
filtered_df = df[df['Date added'] >= last_month]

# Sort by 'Date added' (newest to oldest) and 'Title' (A-Z)
filtered_df = filtered_df.sort_values(by=['Date added', 'Title'], ascending=[False, True])

# Save the filtered and sorted data to a new CSV
output_csv = "siteimprove-sheets/parsed_siteimprove_export.csv"
filtered_df.to_csv(output_csv, index=False)

# Print the type and some samples of successful date conversions
print("Date added column data type:", df['Date added'].dtype)
print("Sample of parsed Date added column:")
print(df['Date added'].head(10))
