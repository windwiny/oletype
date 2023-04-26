import random

import win32com.client
from oletype import excel

exapp: excel.Application = None
exapp = win32com.client.Dispatch('excel.application')
exapp.Visible = True

vv:excel.Workbook = exapp.Workbooks.Add()
print('vv', vv)
vv.HasMailer

wb: excel._Workbook = exapp.ActiveWorkbook
ws = wb.ActiveSheet

print('ws',excel.Workbooks)

print(type(wb))


print(ws.Name)
ws.Name = 'aste 3'
print(ws.Name)

rs = 'A3:B6'

r: excel.Range = ws.Range(rs)
r.Value = [(random.random(), random.random()),
           (random.random(), random.random()),
           (random.random(), random.random()),
           (random.random(), random.random()),
           ]

r.Select()
ws.Shapes.AddChart()

print(ws.Range(rs).Value)
wb.Saved = True
