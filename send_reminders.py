import xlrd
import ezgmail
import os.path

while True:
    file_path = input("Enter the name of the file you wish to read from (Type Q to quit): ")
    file_exists = os.path.isfile(file_path)

    if file_path in ["Q", "q"]:
        print("Terminating program.")
        quit()

    if file_exists:
        break
    else:
        print("Invalid file name. Please try again")

wb = xlrd.open_workbook(file_path)
sheet = wb.sheet_by_index(0)

email_list = []

header = []
for col in range(sheet.ncols):
    header.append(sheet.cell_value(0, col))

list_of_rows = []
for row in range(1, sheet.nrows):
    rows = {}
    for col in range(sheet.ncols):
        rows[header[col]] = sheet.cell_value(row, col)
    list_of_rows.append(rows)

for i in list_of_rows:
    if i['Paid (Y/N)'] == 'N':
        email_list.append(i['Email'])

subject_line = 'Reminder to pay Chanda'
email_body = 'In a few hours, we are coming to an end of our fiscal year. \n\nI request all of you who have not yet ' \
             'completely paid their Khuddam Chanda, or not paid as per the rate, to kindly pay it in full online at ' \
             'chanda.mkausa.org. \n\nWe should always remember that financial contribution is a means of bringing us ' \
             'closer to Allah the Almighty, who provides for us every day of the year, every year. ' \
             '\nMay Allah provide all the means for your physical and spiritual sustenance. JazakAllah. \n\nWassalam,' \
             '\nAreeb Amjad'

print("\nThis is the list of people who have not paid their dues:")
print(email_list)

while True:
    send_email = input("\nDo you want to send a reminder email to these addresses? (Y/N): ")
    if send_email in ["Y", "y"]:
        for email in email_list:
            ezgmail.send(email, subject_line, email_body)
        break
    elif send_email in ["N", "n"]:
        print("Ok. The email won't be sent. Terminating program.")
        break
    else:
        print("Please type either 'Y' (for yes) or 'N' (for no).")

quit()






