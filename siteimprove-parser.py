import pandas as pd
from dateutil import parser

# Load the Siteimprove CSV with UTF-16 encoding, skipping the first three rows
input_csv = "siteimprove-sheets/siteimprove_export.csv"
col_names = ['Title', 'URL', 'Tags', 'Date added', 'Accessibility score', 'Issues']
df = pd.read_csv(input_csv, encoding='utf-16', skiprows=3, delimiter='\t', names=col_names)

# Convert all values in 'Date added' to string to ensure compatibility with .str methods
df['Date added'] = df['Date added'].astype(str)

# Clean up date strings by stripping whitespace and normalizing spaces
df['Date added'] = df['Date added'].str.strip()
df['Date added'] = df['Date added'].str.replace(r'\s+', ' ', regex=True)

# Manually parse the dates using dateutil.parser.parse
def parse_date(date_str):
    try:
        return parser.parse(date_str)
    except parser.ParserError:
        return pd.NaT

df['Date added'] = df['Date added'].apply(parse_date)

# Identify rows where the conversion failed
invalid_dates = df[df['Date added'].isna()]
print("Rows with invalid date format after coercion:")
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

print(df['Date added'].dtype)
print(df['Date added'].head(10))  # Check the first 10 entries
