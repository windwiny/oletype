import sys
sys.path.insert(0, '.')

import time
from win32com import client
from oletype import excel
print(repr(excel))

if __name__ == '__main__':
    print('---')
    eee : excel.Application = client.Dispatch('excel.application')
    eee.Visible = True
    eee.Workbooks.Add()
    cc=eee.ActiveChart
    cd=eee.ActiveSheet
    rg=cd.UsedRange
    rg.Value2
    rg.AddComment(f'test  {time.time()}')
