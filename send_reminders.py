import xlrd

file_path = "Member Dues.xlsx"
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










#
# # First store members' paid statuses in a list
#
# for i in range(1, sheet.nrows):
#     if sheet.cell_value(i, 5) == 'Y':
#         has_paid_list.append(True)
#     if sheet.cell_value(i, 5) == 'N':
#         has_paid_list.append(False)
#
# for i, j in enumerate(has_paid_list):
#     if not j:
#         email_list.append(sheet.cell_value(i + 1, 4))





