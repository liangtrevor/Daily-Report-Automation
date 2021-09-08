# main file
# version 2.6.2 of OpenPyXL
# to install: >>> pip install openpyxl==2.6.2

import openpyxl
from datetime import date
from functions import passthrough

today = date.today()

weekdays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']

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

month_num = today.month

day = today.strftime("%d ")

year = today.year

start = int(input("Enter day to start at: "))

end = int(input("Enter last day: "))

for z in range(start, end + 1):
    # if today.weekday()
    # daywk = days[str(abs(today.weekday() - (7 - z)))]
    week_num = date(int(year),int(month_num),int(day)).weekday()
    weekday = weekdays[week_num - (7 - z)]
    date = str(weekday) + ", " + str(int(day) - (7 + z)) + " " + month # day of the week, day, month
    print("the day of week: " + str(today.weekday()))
    print("Date: " + date)
    tempHigh = input("Enter temperature high: ")
    precip = input("Enter precipitation: ")
    passthrough(pos_report=wb_pos, report=wb_report, precip=precip,
                temphigh=tempHigh, passnumber=z, date=date)

wb_report.save("report_completed.xlsx")

input()