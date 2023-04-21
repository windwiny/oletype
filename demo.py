import random

import win32com.client
import oletype

exapp: oletype.Application = None
exapp = win32com.client.Dispatch('excel.application')
exapp.Visible = True

exapp.Workbooks.Add()
wb: oletype._Workbook = exapp.ActiveWorkbook
ws = wb.ActiveSheet


print(type(wb))


print(ws.Name)
ws.Name = 'aste 3'
print(ws.Name)

rs = 'A3:B6'

r: oletype.Range = ws.Range(rs)
r.Value = [(random.random(), random.random()),
           (random.random(), random.random()),
           (random.random(), random.random()),
           (random.random(), random.random()),
           ]

r.Select()
ws.Shapes.AddChart()

print(ws.Range(rs).Value)
wb.Saved = True
