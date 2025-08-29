import subprocess
import pandas as pd

def run_script(script_path, *args):
    command = ['python', script_path, *args]
    subprocess.run(command, check=True)

# Run the first script
run_script('generate_hosts_01192024.py', 'CSAM-org-acronym.xlsx', 'TestData-Ver2.xlsx', '.')

# Run the second script
run_script('clear_tuple_generate_hosts_new_template.py', 'CSAM-org-acronym.xlsx', 'CombinedFile-NewTemplate.xlsx', '.')

# Run the third script
run_script('updated_generate_hosts.py', 'CSAM-org-acronym.xlsx', 'CombinedFile-OldTemplate.xlsx', '.')


# Initialize an empty list to store DataFrames
dfs = []

# Read the content of each output CSV file and append to the list
df1 = pd.read_csv('generate_hosts.csv')
dfs.append(df1)

df2 = pd.read_csv('clear_tuple_host_new_template.csv', skiprows=1)
dfs.append(df2)

df3 = pd.read_csv('host_data.csv', skiprows=1)
dfs.append(df3)

# Concatenate the list of DataFrames vertically
concatenated_df = pd.concat(dfs, ignore_index=True)

# Save the final CSV file
concatenated_df.to_csv('master_output.csv', index=False)