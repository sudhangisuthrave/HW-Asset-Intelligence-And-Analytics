import openpyxl

# Define the values to check for in any cell of the first row
values_to_check = ["IP Address (Internal)", "NAT IPs", "Identifier or Host Name", "CPU Core"]

# Open the Excel workbook
workbook2 = openpyxl.load_workbook('Clean-CombinedNewTemplate.xlsx')

# Iterate through each sheet in the workbook
for sheet_name in workbook2.sheetnames:
    sheet = workbook2[sheet_name]

    found_values = []

    # Iterate through the cells in the first row of the sheet
    for cell in sheet[1]:
        if cell.value in values_to_check:
            found_values.append(cell.value)

    if found_values:
        print(f"Sheet '{sheet_name}' Found")
    else:
        print(f"Sheet '{sheet_name}' does not have any of the expected values in the first row.")

# Close the workbook
workbook2.close()
