
import win32com.client
from . import excel

exapp: excel.Application = win32com.client.Dispatch('excel.application')

exapp.Visible = True

et: excel.Workbook = exapp.Workbooks.Add()

wb: excel._Workbook = exapp.ActiveWorkbook


ws = wb.ActiveSheet
ws.Buttons

ws.Name = 'test1'

r: excel.Range = ws.Range('A1:b2')

r.Value = [
    (1, 2,),
    (3, 4),
]

ws.Select(r)

ws.Shapes.AddChart2()

wb.Saved = True


v: str | None
v = None


def tk() -> str | None:
    pass


class Double:
    pass


def nnn(a: Double) -> str:
    pass


nnn(11)
