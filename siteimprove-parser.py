import pandas as pd
from dateutil import parser

# Load the Siteimprove CSV with the correct columns
input_csv = "siteimprove-sheets/siteimprove_export.csv"
col_names = ['Title', 'URL', 'Tags', 'Date added', 'Accessibility score', 'Site target', 'Points to target', 'Pages', 'Issues', 'Potential issues', 'PDFs with issues']
df = pd.read_csv(input_csv, encoding='utf-16', skiprows=3, delimiter='\t', names=col_names)

# Select only the necessary columns
df = df[['Title', 'URL', 'Tags', 'Date added', 'Accessibility score', 'Issues']]

# Remove rows where 'Date added' is not a valid date string (i.e., it's the header row or corrupted)
df = df[df['Date added'] != 'Date added']

# Convert 'Date added' to datetime, handling any parsing errors by setting invalid entries to NaT
df['Date added'] = df['Date added'].apply(lambda x: parser.parse(str(x)) if pd.notnull(x) else pd.NaT)

# Ensure the 'Accessibility score' is numeric
df['Accessibility score'] = pd.to_numeric(df['Accessibility score'], errors='coerce')

# Sort the DataFrame by 'Title' alphabetically
df = df.sort_values(by='Title', ascending=True)

# Save the sorted data to a new CSV
output_csv = "siteimprove-sheets/parsed_siteimprove_export.csv"
df.to_csv(output_csv, index=False)

print("Parsed and sorted CSV has been saved.")
