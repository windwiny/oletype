from io import StringIO
import sys
import inspect
import win32com.client

exapp = win32com.client.Dispatch('excel.application')
exapp.Visible = True

wb = exapp.Workbooks.Add()
ws = exapp.ActiveSheet
ws.Range("A1:B3").Value = [[2, 4], [5, 6], [4, 6]]
ws.Shapes.AddChart()

# ------------------------------------------------------------------------

all_cls: dict[str, any] = dict()
already_pr: set[str] = set()
autoclsn: str = "<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9."
autoclsn2: str = "win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9."
mypkgname = 'excel.'
mypkgname = ''


def conv2cls(ff,
             ex,
             cls_name: str,
             o_attrs: list[tuple],
             o_methods: list[tuple],
             o_unknows: list[tuple],
             e_noattr: list[tuple],
             e_ee: list[tuple]
             ):

    print(file=ff)
    print(f'class {cls_name}:', file=ff)

    print('  def __init__(self):', file=ff)
    for i in o_attrs:
        na, ty = i
        ty = str(ty)
        ty = ty.replace("<class '", '').replace(
            autoclsn2, mypkgname).replace("'>", '')
        print(f"    self.{na}: {ty}", file=ff)
    print(file=ff)

    for i in o_methods:
        na, vv = i
        s = inspect.signature(vv)
        ars = []
        for i, j in s.parameters.items():
            ars.append(f"{i}")  # :{i}
        ars = ", ".join(ars)
        ret = s.return_annotation
        ret = str(ret)
        ret = ret.replace("<class '", '').replace("'>", '')
        ret = 'None' if ret == 'inspect._empty' else ret
        print(
            f"  def {na}(self{(', ' + ars) if ars else ''}){' -> '+ret if ret!='None' else ''}:  pass", file=ff)
    print(file=ff)

    print(f'  #unknow:', file=ff)
    for i in o_unknows:
        na, ee = i
        print(f"    # {na}:  {ee}", file=ff)
    print(file=ff)

    print(f"  #getattr AttributeError:", file=ff)
    for i in e_noattr:
        na, ee = i
        print(f"    # {na}:  {ee}", file=ff)
    print(file=ff)

    print(f"  #getattr Exception:", file=ff)
    for i in e_ee:
        na, ee = i
        print(f"    # {na}:  {ee}", file=ff)
    print(file=ff)
    print(f'# Summary "{ex.__class__.__mro__}", attrs:{len(o_attrs)}, methods:{len(o_methods)}, whats:{len(o_unknows)},   ok:{len(o_attrs) + len(o_methods) + len(o_unknows)}, er:{len(e_noattr)}, er:{ len(e_ee)}', file=ff)


def showinfo(ff, cls_name: str, ex):
    if cls_name in already_pr:
        # print(f'### SKIP pred {cls_name}\n', file=sys.stderr)
        return
    else:
        already_pr.add(cls_name)
        # print(f'\n### SHOWINFO {cls_name}', file=sys.stderr)

    objxx = [repr(i) for i in dir(object())]
    xx = dir(ex)
    o_attrs = []
    o_methods = []
    o_unknows = []
    e_noattr = []
    e_ee = []
    for i in xx:
        if repr(i) in objxx:
            continue
        try:
            # ex.__getattr__(i)
            vv = getattr(ex, i)
            typevv = type(vv)
            if typevv in [bool, str, int, float, dict, tuple]:
                o_attrs.append((i, typevv))
            elif str(typevv) == "<class 'method'>":
                o_methods.append((i, vv))
            elif str(typevv).startswith(autoclsn):
                clsn = str(typevv).replace(autoclsn, '').replace("'>", '')
                all_cls[clsn] = vv
                o_attrs.append((i, typevv))
            else:
                o_unknows.append((i, typevv))
        except AttributeError:
            e_noattr.append(i)
        except Exception as e:
            e_ee.append((i, e))

    conv2cls(ff, ex, cls_name, o_attrs, o_methods, o_unknows, e_noattr, e_ee)

    # print("\n".join([i[0] for i in o_unknows]))


cls_name = 'Application'
all_cls[cls_name] = exapp


ff = StringIO()
print(f"# just stub file, import it , declare app obj, get ide auto type hit\n", file=ff)
print(f"import os", file=ff)
print(f"import sys", file=ff)
print(f"import datetime", file=ff)
print(file=ff)
for _ in range(10000):
    if all_cls.keys() == already_pr:
        print("## ALL DONE", file=sys.stderr)
        break
    for clsn in list(all_cls):
        exx = all_cls[clsn]
        showinfo(ff, clsn, exx)
print(f'## PRINT  {len(all_cls)}\n', file=sys.stderr)


data = ff.getvalue()

hasbaseclss = list(
    filter(lambda x: True if x[0] == '_' else False, all_cls.keys()))

for _base in hasbaseclss:
    cls = _base[1:]
    data = data.replace(f'class {cls}:', f'class {cls}({_base}):')

print(data)

print('all cls', len(all_cls), sorted(all_cls.keys()), file=sys.stderr)
print('base cls', len(hasbaseclss), sorted(hasbaseclss), file=sys.stderr)

ws.Name = 'test'
wb.Saved = True
