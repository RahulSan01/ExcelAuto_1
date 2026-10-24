from openpyxl import Workbook
from win32com.client import Dispatch

workbook = Workbook()
sheet = workbook.active

sheet['A1'] = "hello"
sheet['B1'] = "there"

workbook.save(r"C:\Users\rahul\OneDrive\Documents\Excel Stuff\hello_there.xlsx")

xl = Dispatch("Excel.Application")
xl.Visible = True

wb = xl.Workbooks.Open(r'C:\Users\rahul\OneDrive\Documents\Excel Stuff\hello_there.xlsx')
