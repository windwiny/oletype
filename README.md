
# What's this

python win32com excel object pyi file


# Install

## Build from source

### Generate pyi file

```shell
git clone https://github.com/windwiny/oletype

cd oletype
python gen_win32com.py > oletype\__init__.py

python demo.py
```

### install package

```shell
pip install .
```

# How to use

use vscode open `demo.py` file, let coding

let var type

```python
import win32com.client
import oletype

exapp: oletype.Application = None
exapp = win32com.client.Dispatch('excel.application')
exapp.Visible = True

exapp.Workbooks.Add()
wb: oletype._Workbook = exapp.ActiveWorkbook
ws = exapp.ActiveSheet


```

input var and dot, vscode will show quick info and method signatures.
