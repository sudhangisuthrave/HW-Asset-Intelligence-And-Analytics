import subprocess
import pandas as pd


def run_script(script_path, *args):
    command = ['python', script_path, *args]
    subprocess.run(command, check=True)



# Read the content of all output CSV files. Change this line to include the new file.
df4 = pd.read_csv('master_output.csv').dropna(how='all')  # Drop rows where all values are NaN
df5 = pd.read_csv('DontMove-aggregated-hw-inventory-587.csv').dropna(how='all')  # Drop rows where all values are NaN
df6 = pd.read_csv('DontMove-generate_hosts_593.csv').dropna(how='all')  # Drop rows where all values are NaN


# Print the number of rows in each DataFrame for diagnostic purposes
print(f'df1 V2.3 CSAM template rows: {len(df1)}')
print(f'df2 V2.0 Template rows: {len(df2)}')
print(f'df3 Old Template rows: {len(df3)}')

df4.reset_index(drop=True, inplace=True)
df5.reset_index(drop=True, inplace=True)
df6.reset_index(drop=True, inplace=True)

# Concatenate vertically
# concatenated_df = pd.concat([df1, df2, df3], axis=0, ignore_index=True)
concatenated_df = pd.concat([df4, df5, df6], axis=0, join='outer')


# Save the final CSV file
concatenated_df.to_csv('final_CDM_Hostnames.csv', index=False)
