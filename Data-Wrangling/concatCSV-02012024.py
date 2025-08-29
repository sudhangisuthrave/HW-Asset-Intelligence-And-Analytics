import subprocess
import pandas as pd


def run_script(script_path, *args):
    command = ['python', script_path, *args]
    subprocess.run(command, check=True)


# Run the first script
print('New Template Ver 2.3 Inventory Extraction Begins.....')
run_script('generate_hosts_01192024.py', 'CSAM-org-acronym.xlsx', 'CorrectedCombinedFile-CSAMTemplate.xlsx', '.')

# Run the second script
print('New Template Ver 2 Inventory Extraction Begins.....')
run_script('clear_tuple_generate_hosts_new_template.py', 'CSAM-org-acronym.xlsx', 'CorrectedCombinedFile-NewTemplate.xlsx', '.')

# Run the third script
print('Old Template Inventory Extraction Begins.....')
run_script('updated_generate_hosts.py', 'CSAM-org-acronym.xlsx', 'Raw-CombinedFile-OldTemplate.xlsx', '.')


# Read the content of all output CSV files
df1 = pd.read_csv('generate_hosts.csv').dropna(how='all')  # Drop rows where all values are NaN
df2 = pd.read_csv('clear_tuple_host_new_template.csv').dropna(how='all')  # Drop rows where all values are NaN
df3 = pd.read_csv('host_data.csv').dropna(how='all')  # Drop rows where all values are NaN


# Print the number of rows in each DataFrame for diagnostic purposes
print(f'df1 V2.3 CSAM template rows: {len(df1)}')
print(f'df2 V2.0 Template rows: {len(df2)}')
print(f'df3 Old Template rows: {len(df3)}')

df1.reset_index(drop=True, inplace=True)
df2.reset_index(drop=True, inplace=True)
df3.reset_index(drop=True, inplace=True)

# Concatenate vertically
# concatenated_df = pd.concat([df1, df2, df3], axis=0, ignore_index=True)
concatenated_df = pd.concat([df1, df2, df3], axis=0, join='outer')


# Save the final CSV file
concatenated_df.to_csv('master_output.csv', index=False)
