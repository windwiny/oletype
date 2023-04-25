from io import StringIO
import os.path
from pprint import pprint
import sys
import inspect
import win32com.client

obj2methods = {}
obj2parameters = {}
obj2unknow = {}

all_see_type = {}
hasfindtype = {}
unparsedtype = {}

api_fn = 'excel.api.txt'
if os.path.exists(api_fn):
    state = None
    row = 0
    with open(api_fn) as ff:
        while True:
            ll = ff.readline()
            if len(ll) == 0:
                print(f'read {row} lines', file=sys.stderr)
                break
            row += 1
            if ll.startswith('ALL Class.Method '):
                state = 'method'
            elif ll.startswith('ALL Class.Parameter'):
                state = 'parameter'
            else:
                xx = ll.split('->', 2)
                if len(xx) != 2:
                    continue
                clm, objt = xx
                cln, member = clm.strip().split('.', 2)
                objt = objt.strip()
                if state == 'method':
                    if cln not in obj2methods:
                        obj2methods[cln] = {}
                    obj2methods[cln][member] = objt
                elif state == 'parameter':
                    if cln not in obj2parameters:
                        obj2parameters[cln] = {}
                    obj2parameters[cln][member] = objt
                else:
                    if cln not in obj2unknow:
                        obj2unknow[cln] = {}
                    obj2unknow[cln][member] = objt

pprint({'obj2methods': obj2methods}, width=200, stream=sys.stderr)
print(file=sys.stderr)
pprint({'obj2parameters': obj2parameters}, width=200, stream=sys.stderr)
print(file=sys.stderr)
pprint({'obj2unknow': obj2unknow}, width=200, stream=sys.stderr)
print(file=sys.stderr)


def get_cln_method_return(cln, member):
    if cln in obj2methods:
        cl = obj2methods[cln]
        if member in cl:
            objtype = cl[member]
            if objtype == '':
                return None
            elif objtype == 'None':
                return 'None'

            un = False
            for cc in ' \t~`!@#$%^&*()+-={}[]\\|;\':"<>,./?':  # no _
                if cc in objtype:
                    un = True
                    objtype = objtype.replace(cc, '_')
            if un:
                unparsedtype[objtype] = 1
                return objtype
            else:
                if objtype not in ['float','bool','list','int','str']:
                    all_see_type[objtype] = 1
                return objtype

    return None


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


def conv2cls(ex,
             cls_name: str,
             o_attrs: list[tuple],
             o_methods: list[tuple],
             o_unknows: list[tuple],
             e_noattr: list[tuple],
             e_ee: list[tuple]
             ):
    ff = StringIO()
    print(file=ff)
    print(f'class {cls_name}:', file=ff)

    print('  def __init__(self):', file=ff)
    for ix in o_attrs:
        na, ty = ix
        ty = str(ty)
        ty = ty.replace("<class '", '').replace(
            autoclsn2, mypkgname).replace("'>", '')
        print(f"    self.{na}: {ty}", file=ff)
    print(file=ff)

    for ix in o_methods:
        na, vv = ix
        s = inspect.signature(vv)
        ars = []
        for v1, v2 in s.parameters.items():
            ars.append(f"{v1}")  # :{v1}
        ars = ", ".join(ars)
        ret = s.return_annotation
        ret = str(ret)
        ret = ret.replace("<class '", '').replace("'>", '')
        if ret == 'inspect._empty':
            ret = get_cln_method_return(cls_name, na)
        else:
            pass
        if ret is not None:
            rets = ' -> ' + str(ret)
        else:
            rets = ''
        print(
            f"  def {na}(self{(', ' + ars) if ars else ''}){rets}:  pass", file=ff)
    print(file=ff)

    print(f'  #unknow:', file=ff)
    for ix in o_unknows:
        na, ee = ix
        print(f"    # {na}:  {ee}", file=ff)
    print(file=ff)

    print(f"  #getattr AttributeError:", file=ff)
    for ix in e_noattr:
        na, ee = ix
        print(f"    # {na}:  {ee}", file=ff)
    print(file=ff)

    print(f"  #getattr Exception:", file=ff)
    for ix in e_ee:
        na, ee = ix
        print(f"    # {na}:  {ee}", file=ff)
    print(file=ff)
    print(f'# Summary "{ex.__class__.__mro__}", attrs:{len(o_attrs)}, methods:{len(o_methods)}, whats:{len(o_unknows)},   ok:{len(o_attrs) + len(o_methods) + len(o_unknows)}, er:{len(e_noattr)}, er:{ len(e_ee)}', file=ff)
    return ff.getvalue()


def showinfo(cls_name: str, ex):
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

    cls_def_str = conv2cls(ex, cls_name, o_attrs,
                           o_methods, o_unknows, e_noattr, e_ee)
    return cls_def_str


cls_name = 'Application'
all_cls[cls_name] = exapp


all_cls_def_strs = []

for _ in range(10000):
    if all_cls.keys() == already_pr:
        print("## ALL DONE", file=sys.stderr)
        break
    for clsn in list(all_cls):
        exx = all_cls[clsn]
        cls_def_str = showinfo(clsn, exx)
        all_cls_def_strs.append(cls_def_str)
        hasfindtype[clsn] = 1


def output_all():
    global unparsedtype, all_cls_def_strs
    global hasfindtype, nofindtype, unparsedtype
    ff = StringIO()

    # head
    print(f"# just stub file, import it , declare app obj, get ide auto type hit\n", file=ff)
    print(f"import os", file=ff)
    print(f"import sys", file=ff)
    print(f"import datetime", file=ff)
    print(file=ff)

    # output all unparsed
    print('# TODO FIXME unparsed', file=ff)
    for i in sorted(unparsedtype):
        print(f'class {i}: pass', file=ff)
    print(file=ff)

    # output all nofind define
    print('# TODO FIXME nofind define', file=ff)
    for i in sorted(all_see_type):
        if i not in hasfindtype:
            print(f'class {i}: pass', file=ff)
    print(file=ff)

    # reverse class define order
    while True:
        if len(all_cls_def_strs) == 0:
            break
        data = all_cls_def_strs.pop()
        print(data, file=ff)

    print(f'## PRINT  {len(all_cls)}\n', file=sys.stderr)

    # get all
    data = ff.getvalue()

    # replace CoCobase
    hasbaseclss = list(
        filter(lambda x: True if x[0] == '_' else False, all_cls.keys()))
    for _base in hasbaseclss:
        cls = _base[1:]
        data = data.replace(f'class {cls}:', f'class {cls}({_base}):')

    # output all define
    print(data)

    print('all cls', len(all_cls), sorted(all_cls.keys()), file=sys.stderr)
    print('base cls', len(hasbaseclss), sorted(hasbaseclss), file=sys.stderr)


output_all()

ws.Name = 'test'
wb.Saved = True
