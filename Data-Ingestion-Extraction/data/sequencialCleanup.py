import subprocess
import pandas as pd


def run_script(script_path, *args):
    command = ['python', script_path, *args]
    subprocess.run(command, check=True)


# Run the first script
print('Combining all HWAM files and remove extra sheets......')
run_script('combineNewTemplate-RemoveExtraSheets.py')

# Run the second script
print('Separate all templates based on versions used, V2.0 or V2.3 or Old Template.......')
run_script('seperateNew-Old-Latest-TemplateHW.py')


# Run the third script
print('Removing unwanted rows from V2.0 inventories.....')
run_script('conditionalDeleteUnwantedRows.py')


# Run the forth script
print('Removing unwanted rows from V2.3 inventories.....')
run_script('v-2-3-conditionalDeleteUnwantedRows.py')


# Run the fifth script
print('Check for missing columns and insert them wherever necessary from V2.3 inventories.....')
run_script('insertMissingColumnsCSAMTemplate.py')


# Run the sixth script
print('Check for missing columns and insert them wherever necessary from V2.0 inventories.....')
run_script('insertMissingColumnsNewTemplate.py')





