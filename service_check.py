import pandas as pd
import os

def load_file(filename):
    """Load Excel or CSV file based on extension."""
    ext = os.path.splitext(filename)[1].lower()
    if ext == '.xlsx':
        return pd.read_excel(filename, engine='openpyxl')
    elif ext == '.csv':
        return pd.read_csv(filename)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

def clean_dataframe(df):
    """Clean dataframe by removing empty rows, whitespace, and duplicates"""
    df = df.dropna(subset=['Name', 'Status'])
    for col in df.select_dtypes(include=['object']):
        df[col] = df[col].str.strip()
    # Remove rows where Name or Status is now empty after stripping
    df = df[(df['Name'] != '') & (df['Status'] != '')]
    return df

# Change these filenames as needed
new_filename = 'new.csv'  # or 'new.xlsx'
old_filename = 'old.csv'  # or 'old.xlsx'

# Load the files (Excel or CSV)
new_file = clean_dataframe(load_file(new_filename))
old_file = clean_dataframe(load_file(old_filename))

# Normalize service names by removing UserID (e.g., _36fdd424)
new_file['NormalizedName'] = new_file['Name'].str.replace(r'_[a-zA-Z0-9]+$', '', regex=True)
old_file['NormalizedName'] = old_file['Name'].str.replace(r'_[a-zA-Z0-9]+$', '', regex=True)

# Remove duplicates based on NormalizedName (keep first occurrence)
new_file = new_file.drop_duplicates(subset=['NormalizedName'])
old_file = old_file.drop_duplicates(subset=['NormalizedName'])

# Set the normalized service name as the index
new_file.set_index('NormalizedName', inplace=True)
old_file.set_index('NormalizedName', inplace=True)

# Find services only in the new file
new_only_services = new_file[~new_file.index.isin(old_file.index)]

# Find services only in the old file
old_only_services = old_file[~old_file.index.isin(new_file.index)]

# Find services in both files with different states
common_services = new_file[new_file.index.isin(old_file.index)]

different_state_services = common_services[
    common_services['Status'] != old_file.loc[common_services.index, 'Status']
]

# Output to a text file
with open('service_comparison_report.txt', 'w', encoding='utf-8') as f:
    f.write("\n=== Service Comparison Report ===\n\n")

    f.write("1. Services only existing this month:\n")
    f.write("-" * 40 + "\n")
    if len(new_only_services) > 0:
        for idx, row in new_only_services.iterrows():
            if idx and row['Status']:  # filter out empty lines
                f.write(f"Service: {idx}\n")
                f.write(f"Status: {row['Status']}\n\n")
    else:
        f.write("No new services found\n\n")

    f.write("2. Services only existing last month:\n")
    f.write("-" * 40 + "\n")
    if len(old_only_services) > 0:
        for idx, row in old_only_services.iterrows():
            if idx and row['Status']:  # filter out empty lines
                f.write(f"Service: {idx}\n")
                f.write(f"Status: {row['Status']}\n\n")
    else:
        f.write("No removed services found\n\n")

    f.write("3. Services with changed states:\n")
    f.write("-" * 40 + "\n")
    if len(different_state_services) > 0:
        for idx, row in different_state_services.iterrows():
            if idx and row['Status']:  # filter out empty lines
                f.write(f"Service: {idx}\n")
                f.write(f"Current Status: {row['Status']}\n")
                f.write(f"Previous Status: {old_file.loc[idx, 'Status']}\n\n")
    else:
        f.write("No services with changed states found\n\n")