import win32com
from win32com.client import Dispatch
from win32com.client import constants
import win32com.client as win32

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = -1
#此处的change文件放在deal路径中报错，因此复制到另外的目录下。地址需为绝对地址
wb = excel.Workbooks.Open('C:\change.xlsx')
ws = wb.Worksheets[0]

m=1410

for row1 in range(2, m):
    for column1 in range(1, 11):
        if ws.cell(row=row1, column=column1).value is None:
            ws.Rows("row1").delete()

wb.Save("change_mid.xlsx")