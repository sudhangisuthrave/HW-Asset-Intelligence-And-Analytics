import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import xml.etree.ElementTree as ET

# Load the Excel file
workbook = openpyxl.load_workbook('Clean-CombinedNewTemplate.xlsx', data_only=True)

# Iterate through all sheets in the workbook
for sheet_name in workbook.sheetnames:
    sheet = workbook[sheet_name]
    xml_source = sheet.sheet._data_source

    # Parse the XML data to identify cells with formulas
    root = ET.fromstring(xml_source)
    for c in root.iter(ET.QName(root.tag).namespace + 'c'):
        formula = c.find('.//' + ET.QName(root.tag).namespace + 'f')
        if formula is not None:
            cell_coordinate = c.get('r')
            print(f"Sheet: {sheet_name}, Cell {cell_coordinate} contains a formula: {formula.text}")

# Close the Excel file when you're done
workbook.close()
