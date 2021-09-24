# main file
# version 2.6.2 of OpenPyXL
# to install: >>> pip install openpyxl==2.6.2

import openpyxl
from datetime import date
from functions import passthrough

today = date.today()

weekdays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']

print("Hi Ange, welcome to the automation wizard!")
print("You will be prompted for a filename")
print("Please include the file extension at the end (e.g. .xlsx) \n")
userResponse = input("Enter the filename of POS report: ")

wb_pos = openpyxl.load_workbook(userResponse)

userResponse = input("Enter the filename of report template: ")

wb_report = openpyxl.load_workbook(userResponse)

# filename for saving new sheet at the end
newFilename = userResponse + "_completed"

month = today.strftime("%b")

month_num = today.month

day = today.strftime("%d ")

year = today.year

start = int(input("Enter day to start at: "))

end = int(input("Enter last day (excluded): "))

dayofweek = today.weekday()

startDate = (end - (end - start))

for z in range(start, end):
    weekday = weekdays[int(dayofweek) - (end - z)]
    date = str(weekday) + ", " + str(int(day) - (end - z)) + " " + month # day of the week, day, month
    print(date)
    startDate += 1
    precip = input("Enter precipitation: ")
    tempHigh = input("Enter temperature high: ")
    passthrough(pos_report=wb_pos, report=wb_report, precip=precip,
                temphigh=tempHigh, passnumber=z, date=date)

wb_report.save("report_completed.xlsx")

print("All sheets up to " + date + " populated.")

input()