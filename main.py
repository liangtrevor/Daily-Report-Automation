# main file
# version 2.6.2 of OpenPyXL
# to install: >>> pip install openpyxl==2.6.2

import openpyxl
from datetime import date
from functions import passthrough

today = date.today()

days = {'0': 'Mon', '1': 'Tue', '2': 'Wed', '3': 'Thu', '4': 'Fri',
        '5': 'Sat', '6': 'Sun'}

print("Hi, welcome to the automation wizard!")
print("You will be prompted for filenames.")
print("Please include the file extension at the end (e.g. .xlsx) \n")
userResponse = input("Enter the filename of POS report: ")

wb_pos = openpyxl.load_workbook(userResponse)

sheet_pos = wb_pos.active

userResponse = input("Enter the filename of report template: ")

wb_report = openpyxl.load_workbook(userResponse)

# filename for saving new sheet at the end
newFilename = userResponse + "_completed"

month = today.strftime("%b")

day = today.strftime("%d ")

userResponse = input("Enter temperature high: ")

tempHigh = userResponse

userResponse = input("Enter precipitation: ")

precip = userResponse

for z in range(1, 8):

    daywk = days[str(today.weekday() + (z - 1))]

    date = str(daywk) + ", " + str(int(day) + (z - 1)) + " " + month
    passthrough(pos_report=wb_pos, report=wb_report, precip=precip,
                temphigh=tempHigh, passnumber=z, date=date)

wb_report.save("report_completed.xlsx")

input()