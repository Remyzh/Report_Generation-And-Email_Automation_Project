from win32com.client import Dispatch


namesFile = open("File_names.txt", "r")
sheetName = open("Names2.txt", "r")


path1 = 'C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test\\source.xlsx'            #path of the source file VARIABLE
path2 = "C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test\\AccountManagers\\Albert Tompson's Q4 FY2021 Results.xlsx"        #path of the destination file VARIABLE

xl = Dispatch("Excel.Application")
xl.Visible = True  # You can remove this line if you don't want the Excel application to be visible


wb1 = xl.Workbooks.Open(Filename=path1)
wb2 = xl.Workbooks.Open(Filename=path2)

ws1 = wb1.Worksheets('Albert Tompson Results')        #put sheetname that you want to copy (the persons name + format) VARIABLE
ws1.Copy(Before=wb2.Worksheets(1))      #set position in destination file you want to paste in. VARIABLE

wb2.Close(SaveChanges=True)
xl.Quit()