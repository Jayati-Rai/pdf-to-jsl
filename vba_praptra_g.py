import openpyxl

# Load the workbook
workbook_path = "backcode.xlsx"  # Replace with the actual path to your workbook
workbook = openpyxl.load_workbook(workbook_path)

# Create a new sheet
new_sheet_name = "NewSheet"  # Replace with your desired sheet name
if new_sheet_name not in workbook.sheetnames:
    workbook.create_sheet(title=new_sheet_name)
    print(f"Sheet '{new_sheet_name}' created successfully.")
else:
    print(f"Sheet '{new_sheet_name}' already exists.")

# Save the workbook
workbook.save(workbook_path)