# What's this

python win32com excel object pyi file

# Install

## Build from source

### Generate pyi file

```shell
git clone https://github.com/windwiny/oletype

cd oletype

# download web api
ruby downapi.rb > excel.api.txt

# python inspect objecty, list win32com objects's methods, and parameters
#  method return type may not show, find from download api
python gen_win32com.py > oletype\excel.pyi

python demo.py
```

### install package

```shell

pip wheel ./
dir *.whl
pip install oletype-0.3.0-py3-none-any.whl

```

# How to use

use vscode open `demo.py` file, let coding

let var type

```python
import win32com.client
from oletype import excel

exapp: excel.Application = None
exapp = win32com.client.Dispatch('excel.application')
exapp.Visible = True

exapp.Workbooks.Add()
wb: excel._Workbook = exapp.ActiveWorkbook
ws = exapp.ActiveSheet


```

input var and dot, vscode will show quick info and method signatures.
