# demo
import sys
sys.path.insert(0, '.')

from win32com.client import Dispatch
from oletype import excel
print(repr(excel))

if __name__ == '__main__':
    ee : excel.Application = Dispatch('excel.application')
    ee.Visible = True

    ee.Workbooks.Add()
    # ws:excel._Worksheet =
    ws = ee.ActiveSheet
    # rg = ws.Select('A3')
    rg = ws.UsedRange
    rg.Value2
    rg.AddComment('asf')


