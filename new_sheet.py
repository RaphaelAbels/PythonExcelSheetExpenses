import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Fill, Font

font = Font(name="Arial", size=11, bold=True, italic=False, vertAlign=None, underline=None, strike=False, color="00000000")


def sum_it_up(row, start, times):
    sum = "=Sum("
    value = ""
    for x in range(start, start + times):
        value += row + str(x) + ","

    sum += value[0:-1] + ")"
    return sum


def monthly_expenses(row, row_end, start, times, sheet):
    sheet[row + str(start)] = topic
    topic_font = sheet[row + str(start)]
    topic_font.font = font
    sheet[row + str(int(start + 1))] = name
    name_font = sheet[row + str(int(start + 1))]
    name_font.font = font
    sheet[row_end + str(start + 1)] = value
    value_font = sheet[row_end + str(start + 1)]
    value_font.font = font

    for space in range(start + times, start + times + 1):
        sheet[row + str(space)] = sum
        sum_font = sheet[row + str(space)]
        sum_font.font = font
        sheet[row_end + str(space)] = sum_it_up(row_end, start, times)
        sum_sum_font = sheet[row_end + str(space)]
        sum_sum_font.font = font


def daily_expenses(count, identifier, date, row, end, start, times, sheet):
    sheet[count + str(start)] = "NUMBER:"
    n1 = sheet[count + str(start)]
    n1.font = font
    sheet[identifier + str(start)] = "SUBJECT:"
    s1 = sheet[identifier + str(start)]
    s1.font = font
    sheet[date + str(start)] = "DATE:"
    n2 = sheet[date + str(start)]
    n2.font = font
    sheet[row + str(start)] = "IDENTIFIER:"
    n3 = sheet[row + str(start)]
    n3.font = font
    sheet[end + str(start)] = "AMOUNT:"
    n4 = sheet[end + str(start)]
    n4.font = font
    sheet[count + str(start + times + 1)] = "SUM:"
    n5 = sheet[count + str(start + times + 1)]
    n5.font = font
    sheet[end + str(start + times + 1)] = sum_it_up(end, start + 1, times)
    n6 = sheet[end + str(start + times + 1)]
    n6.font = font
    counter = 1
    for x in range(start + 1, start + times + 1):
        sheet[count + str(x)] = counter
        xy = sheet[count + str(x)]
        xy.font = font
        counter = counter + 1


workbook = openpyxl.Workbook()
sheet_active = workbook.active
sheet_active.title = "RENAME"
topic = "TOPIC:"
name = "NAME:"
value = "VALUE:"
sum = "SUM:"
expenses = "EXPENSES:"
sheet_active["A1"] = "This sheet is for monthly calculations to get a better overview of expenses"
a1 = sheet_active["A1"]
a1.font = font
sheet_active["A3"] = "MONTHLY ASSETS:"
a3 = sheet_active["A3"]
a3.font = font
sheet_active["D3"] = "PUT ASSETS HERE"
d3 = sheet_active["D3"]
d3.font = font
sheet_active["A4"] = "CURRENT BALANCE:"
a4 = sheet_active["A4"]
a4.font = font
sheet_active["A6"] = "CURRENT EXPENSES:"
a6 = sheet_active["A6"]
a6.font = font
sheet_active["D4"] = "=D3-D19-I19-D31-I31-D43-I43-P55"
d4 = sheet_active["D4"]
d4.font = font
sheet_active["D6"] = "=SUM(D19,I19,D31,I31,D43,I43,D55,I55,P55)"
d6 = sheet_active["D6"]
d6.font = font
sheet_active["A8"] = "MONTHLY EXPENSES"
a8 = sheet_active["A8"]
a8.font = font
sheet_active["L8"] = "DAILY EXPENSES"
l8 = sheet_active["L8"]
l8.font = font
monthly_expenses("A", "D", 9, 10, sheet_active)
monthly_expenses("F", "I", 9, 10, sheet_active)
monthly_expenses("A", "D", 21, 10, sheet_active)
monthly_expenses("F", "I", 21, 10, sheet_active)
monthly_expenses("A", "D", 33, 10, sheet_active)
monthly_expenses("F", "I", 33, 10, sheet_active)
monthly_expenses("A", "D", 45, 10, sheet_active)
monthly_expenses("F", "I", 45, 10, sheet_active)
daily_expenses("K", "L",  "M", "N", "Q", 9, 45, sheet_active)



#monthly_expenses("K","N",9,10)
#monthly_expenses("P","S",9,10)

#monthly_expenses("K","N",21,10)
#monthly_expenses("P","S",21,10)

workbook.save("Montly Calculations.xlsx")
