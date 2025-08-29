import openpyxl

# Define the workbook name and the list of expected column names
workbook_name = 'CombinedFile-NewTemplate.xlsx'
expected_columns = ['IP Address (Internal)', 'IP Address (External)', 'NAT IPs', 'Identifier or Host Name',	'AD Domain', 'CPU Core', 'Memory (GB)',	'Drive Space (GB)',	'OS Name',	'OS Version', 'Lifecycle',	'Location', 'Hosting/CSP Contract', 'Asset Category', 'Asset Type', 'Virtual',	'Hardware Make',	'Hardware Model',	'Manufacturer Serial Number',	'BIOS UUID/GUID', 'MAC Address(es)', 'Public',	'High Value Asset',	'GFE',	'Date Device Added to System Boundary',	'System Owner / Device Manager', 'Device Operator',	'Systems Supported']  # Replace with your actual column names


def check_columns_in_worksheet(worksheet, expected_columns):
    """
    Check if the first row of the worksheet contains all the expected columns.
    """
    # Get the first row values
    first_row = [cell.value for cell in worksheet[1]]

    # Check if all expected columns are in the first row
    missing_columns = [col for col in expected_columns if col not in first_row]

    return missing_columns


def main():
    # Load the workbook
    wb = openpyxl.load_workbook(workbook_name)

    # Iterate through all the worksheets in the workbook
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        missing_columns = check_columns_in_worksheet(ws, expected_columns)

        if missing_columns:
            print(f"Worksheet '{sheet_name}' is missing columns: {missing_columns}")
        else:
            print(f"Worksheet '{sheet_name}' contains all the expected columns.")


if __name__ == "__main__":
    main()
