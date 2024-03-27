import pandas as pd
from datetime import datetime, timedelta
import win32com.client as win32

# Get the current date and the Monday date of the current week
current_date = datetime.now()
monday_date = current_date - timedelta(days=current_date.weekday())
formatted_monday_date = monday_date.strftime("%Y-%m-%d")
# Formulate the file path based on the Monday date
file_path = f'C:/Users/Ryan.Sissom/Documents/Craddle-2-Grave-{formatted_monday_date}.xlsx'
# Load the Excel file
excel_file = pd.ExcelFile(file_path)
# Get the current date as a string
current_date_str = datetime.now().strftime("%Y-%m-%d")

# Formulate the output file paths including the Monday date
output_file_for_donna_path = f'C:/Users/Ryan.Sissom/Documents/for_donna_{formatted_monday_date}.xlsx'
output_file_for_deepak_path = f'C:/Users/Ryan.Sissom/Documents/for_deepak_{formatted_monday_date}.xlsx'

output_file_for_donna = pd.ExcelWriter(output_file_for_donna_path, engine='xlsxwriter')
output_file_for_deepak = pd.ExcelWriter(output_file_for_deepak_path, engine='xlsxwriter')

# Iterate through each sheet in the Excel file
for sheet_name in excel_file.sheet_names:
    # Read the sheet into a DataFrame
    sheet_df = excel_file.parse(sheet_name)

    # Filter based on the condition qtyord == qtyship
    for_donna_df = sheet_df[sheet_df['qtyord'] == sheet_df['qtyship']]
    # Filter based on the condition qtyord != qtyship
    for_deepak_df = sheet_df[sheet_df['qtyord'] != sheet_df['qtyship']]

    # Write the filtered DataFrames to the respective output files
    for_donna_df.to_excel(output_file_for_donna, sheet_name, index=False)
    for_deepak_df.to_excel(output_file_for_deepak, sheet_name, index=False)

# Save the changes to the output files
output_file_for_donna.close()
output_file_for_deepak.close()

# Outlook email settings
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.Subject = 'Output Files'
mail.Body = 'Please find attached the output files.'

# Attach the files to the email
mail.Attachments.Add(output_file_for_donna_path)
mail.Attachments.Add(output_file_for_deepak_path)

# Email recipients (replace with actual email addresses)
mail.To = 'eli.romero@pssigroup.com'

# Send the email
mail.Send()
