from openpyxl import Workbook
import csv

csv_rowlist = [["Member ID", "Last Name", "First Name", "Phone Number", "Email", "Paid (Y/N)",
                "Pledged Amount", "Paid Amount"],
               ["45372", "Care", "Obama", "775-849-7315", "obamacare@obama.gov", "Y", "$100", "$100"],
               ["688975", "Noor", "Sisay", "972-556-7048", "snoor3883@gmail.com", "N", "$100", "$0"],
               ["405629", "Chibuike", "Sibongile", "773-333-3114", "sibchi1100@yahoo.com", "Y", "$150", "$150"],
               ["43658", "Sikandar", "Abdul", "619-863-1318", "avdol3217@gmail.com", "N", "$95", "$0"],
               ["597757", "Tariku", "Idir", "215-249-7932", "itariku373@gmail.com", "N", "$135", "$0"],
               ]
wb = Workbook()
ws = wb.active
with open('Member Dues List.csv', mode='w', newline='') as file:
    writer = csv.writer(file)
    writer.writerows(csv_rowlist)

with open('Member Dues List.csv', mode='r') as file:
    for row in csv.reader(file):
        ws.append(row)
wb.save('Member Dues.xlsx')

