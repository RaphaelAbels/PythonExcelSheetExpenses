import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Fill

#font = Font(name="Arial", size=11, bold=False, )
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Current Month"

name = "NAME:"
value = "VALUE:"
sum = "SUM:"
expenses = "EXPENSES:"

sheet["A1"] = "This sheet is for usage for monthly calculations to get a better overview of expenses"
sheet["A4"] = "Regular Monthly expenses"
sheet["A5"] = "TOPIC:"
sheet["A7"] = name
sheet["D7"] = value
sheet["A15"] = expenses
sheet["D15"] = "=SUM(D8,D9,D10,D11,D12,D13,D14)"

sheet["F5"] = "MONTHLY ASSETS"
sheet["I5"] = value
sheet["F8"] = "CURRENT BALANCE"
sheet["I8"] = "=I5-D15-D49"

sheet["A17"] = "CURRENT EXPENSES"
sheet["A19"] = name
sheet["D19"] = value
sheet["A49"] = sum
sheet[
    "D49"] = "=Sum(D20,D21,D22,D23,D24,D25,D26,D27,D28,D29,D30,D31,D32,D33,D34,D35,D36,D37,D38,D39,D40,D41,D42,D43,D44,D45,D46,D47,D48)"

workbook.save("Calculator.xlsx")
