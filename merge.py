import pandas as pd
import glob

# List all Excel files in the current directory
excel_files = glob.glob('ids/*.xlsx')

# Create an empty list to store dataframes
dfs = []

# Read each Excel file and append its dataframe to the list
for file in excel_files:
    df = pd.read_excel(file, engine='openpyxl')
    dfs.append(df)

# Merge all dataframes into a single dataframe
merged_df = pd.concat(dfs, ignore_index=True)

# Split merged dataframe into chunks of 999,999 rows
chunk_size = 999998
chunks = [merged_df[i:i+chunk_size] for i in range(0, len(merged_df), chunk_size)]

# Save each chunk to a new sheet in the output Excel file
with pd.ExcelWriter('ids/ids.xlsx', engine='openpyxl') as writer:
    for i, chunk in enumerate(chunks):
        chunk.to_excel(writer, sheet_name=f'Sheet{i+1}', index=False)
