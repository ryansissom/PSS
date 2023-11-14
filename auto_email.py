import pandas as pd
from datetime import datetime
import pandas as pd
import win32com.client as win32
import os

current_date = datetime.now().strftime("%m.%d.%Y")
print(current_date)

# Read the Excel file into a DataFrame
xls = pd.ExcelFile('C:/Users/Ryan.Sissom/Documents/c2gTEST.xlsx')

# Define the output file paths
equal_output_file = 'C:/Users/Ryan.Sissom/Documents/for_donna.xlsx'
not_equal_output_file = 'C:/Users/Ryan.Sissom/Documents/not_equal_output.xlsx'

# Create Excel writers with openpyxl engine for the output files
with pd.ExcelWriter(equal_output_file, engine='openpyxl') as equal_writer, \
        pd.ExcelWriter(not_equal_output_file, engine='openpyxl') as not_equal_writer:
    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name)
        df = df.drop_duplicates()
        equal_df = df[df['qtyord'] == df['qtyship']]
        not_equal_df = df[df['qtyord'] != df['qtyship']]

        # Save the equal DataFrame to the equal output file with the same sheet name
        equal_df.to_excel(equal_writer, sheet_name=sheet_name, index=False)

        # Save the not equal DataFrame to the not equal output file with the same sheet name
        not_equal_df.to_excel(not_equal_writer, sheet_name=sheet_name, index=False)

    equal_writer.book.save(filename=equal_output_file)
    not_equal_writer.book.save(filename=not_equal_output_file)


for_donna_attachment_path = 'C:/Users/Ryan.Sissom/Documents/for_donna.xlsx'
not_equal_output_attachment_path = 'C:/Users/Ryan.Sissom/Documents/not_equal_output.xlsx'

# Create an Outlook application instance
outlook = win32.gencache.EnsureDispatch('Outlook.Application')

# Create and send an email for 'for_donna.xlsx'
for_donna_mail = outlook.CreateItem(0)
for_donna_mail.Subject = 'Attachment for Donna'
for_donna_mail.Body = 'Please see the attached file for Donna.'
for_donna_attachment = os.path.abspath(for_donna_attachment_path)
for_donna_mail.Attachments.Add(for_donna_attachment)
for_donna_mail.To = 'robert.oliver@pssigroup.com'
for_donna_mail.Send()

# Create and send an email for 'not_equal_output.xlsx' to a different recipient
not_equal_output_mail = outlook.CreateItem(0)
not_equal_output_mail.Subject = 'Attachment for a Deepak'
not_equal_output_mail.Body = 'Please see attached file for Deepak.'
not_equal_output_attachment = os.path.abspath(not_equal_output_attachment_path)
not_equal_output_mail.Attachments.Add(not_equal_output_attachment)
not_equal_output_mail.To = 'robert.oliver@pssigroup.com'
not_equal_output_mail.Send()
