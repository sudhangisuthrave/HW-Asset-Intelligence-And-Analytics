import pandas as pd
from lxml import etree


def is_valid_xml(xml_string):
    try:
        etree.fromstring(xml_string)
        return True
    except etree.XMLSyntaxError:
        return False


def is_valid_value(value):
    return isinstance(value, (str, int, float))


def get_invalid_cell_info(df):
    invalid_cells = []

    # Iterate through each cell in the DataFrame
    for index, row in df.iterrows():
        for col_idx, cell_value in enumerate(row):
            # Check if the cell contains a valid value (string or numerical)
            if not is_valid_value(cell_value):
                invalid_cells.append({
                    'row': index + 1,
                    'column': col_idx + 1,
                    'value': cell_value,
                    'reason': 'Invalid value' if not pd.isna(cell_value) else 'Empty cell'
                })

    return invalid_cells


def can_copy_to_another_workbook(file_path):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path, header=None)

    # Get information about invalid cells
    invalid_cells = get_invalid_cell_info(df)

    # Check if there are any invalid cells
    if not invalid_cells:
        print("All cells contain valid values.")

        # Now you can copy the DataFrame to another workbook if needed
        # For example, you can use df.to_excel() to save it to a new workbook

        # Uncomment the following line to save the DataFrame to another workbook
        # df.to_excel('output_workbook.xlsx', index=False)
    else:
        print("Invalid values found:")
        for cell_info in invalid_cells:
            print(f"Row {cell_info['row']}, Column {cell_info['column']}: {cell_info['value']} - {cell_info['reason']}")


# Replace 'your_file.xlsx' with the actual path to your Excel file
file_path = 'C:\\Users\\Sudhangi.Suthrave\\PycharmProjects\\inventory\\csam_inventory\\data\\hwam-209.xlsx'

can_copy_to_another_workbook(file_path)
