from io import StringIO
import json
import os.path
from pprint import pprint
import sys
import inspect
from win32com.client import CDispatch, Dispatch

# TODO


class Cfg():
    def __init__(self) -> None:
        self.num = 0

        self.api_fn = 'excel.api.txt'
        self.api_comment_fn = 'excel.apicomment.json'
        self.out_pyfile_fn = 'oletype/excel.py'

        self.mypkgname = 'excel.'
        self.mypkgname = ''
        # python \Python311\Lib\site-packages\win32com\client\makepy.py
        #  Generating to C:\Users\user1\AppData\Local\Temp\gen_py\3.11\00020813-0000-0000-C000-000000000046x0x1x9.py
        # self.autoclsn: str = "<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9."
        # self.autoclsn2: str = "win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9."
        self.win32comclsn: str = "<class 'win32com."


def check_or_exit(cfg: Cfg):
    if not os.path.exists(cfg.api_fn):
        print(f'not exists api file {cfg.api_fn}')
        sys.exit(1)
    if not os.path.exists(cfg.api_comment_fn):
        print(f'not exists method comment file {cfg.api_comment_fn}')
        sys.exit(2)


def load_comments(fn, clsn_member_comment_kvs):
    with open(fn, encoding='utf-8', errors='ignore') as ff:
        dd = ff.read()
        clsn_member_comment_kvs.update(json.loads(dd))
    print(f'''from {fn} read {len(dd)} bytes, load {len(clsn_member_comment_kvs)} methods comments.''', file=sys.stderr)


def load_apis(fn, obj2methods, obj2parameters, obj2unknow):
    state = None
    row = 0
    with open(fn, encoding='utf-8', errors='ignore') as ff:
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

    pprint({'obj2methods': len(obj2methods)}, width=200, stream=sys.stderr)
    print(file=sys.stderr)
    pprint({'obj2parameters': len(obj2parameters)}, width=200, stream=sys.stderr)
    print(file=sys.stderr)
    pprint({'obj2unknow': len(obj2unknow)}, width=200, stream=sys.stderr)
    print(file=sys.stderr)


def get_cln_method_comment(cln, member, clsn_member_comment_kvs) -> str:
    k = f'{cln}_{member}'
    comment = clsn_member_comment_kvs.get(k)
    return comment


def get_cln_method_return(cln: str, member: str, all_see_type, unparsed_type, obj2methods) -> str | None:
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
                unparsed_type[objtype] = 1
                return objtype
            else:
                if objtype not in ['float', 'bool', 'list', 'int', 'str']:
                    all_see_type[objtype] = 1
                return objtype

    return None


def conv2cls(ex: CDispatch,
             cls_name: str,
             o_attrs: list[tuple],
             o_methods: list[tuple],
             o_unknowns: list[tuple],
             e_noattr: list[tuple],
             e_ee: list[tuple],
             cfg: Cfg,
             all_see_type,
             unparsed_type,
             obj2methods,
             clsn_member_comment_kvs,
             ) -> str:
    '''from dict find cls attrs, methods, unknown properties, and errors'''
    cfg.num += 1

    ff = StringIO()
    print(file=ff)
    print(f'# num={cfg.num}', file=ff)
    print(f'class {cls_name}:', file=ff)

    print('  def __init__(self):', file=ff)
    for ix in o_attrs:
        na, ty = ix
        ty = str(ty)
        ty = cfg.mypkgname + ty.replace("<class '", '').replace("'>", '').split('.')[-1]
        comment_or_pass = get_cln_method_comment(cls_name, na, clsn_member_comment_kvs)
        comment_or_pass = f"\n    '''{comment_or_pass}'''" if comment_or_pass else ''
        print(f"    self.{na}: {ty}{comment_or_pass}", file=ff)
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
            ret = get_cln_method_return(cls_name, na, all_see_type, unparsed_type, obj2methods)
        else:
            pass
        if ret is not None:
            rets = ' -> ' + str(ret)
        else:
            rets = ''
        comment_or_pass = get_cln_method_comment(cls_name, na, clsn_member_comment_kvs)
        comment_or_pass = f"\n    '''{comment_or_pass}'''" if comment_or_pass else '  pass'
        print(f"  def {na}(self{(', ' + ars) if ars else ''}){rets}:{comment_or_pass}", file=ff)
    print(file=ff)

    print(f'  #unknown:', file=ff)
    for ix in o_unknowns:
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
    print(f'# Summary "{ex.__class__.__mro__}", attrs:{len(o_attrs)}, methods:{len(o_methods)}, whats:{len(o_unknowns)},   ok:{len(o_attrs) + len(o_methods) + len(o_unknowns)}, er:{len(e_noattr)}, er:{ len(e_ee)}', file=ff)
    return ff.getvalue()


def showinfo(cls_name: str, ex: CDispatch, all_cls: dict, already_pr: set, all_see_type, unparsed_type, obj2methods, clsn_member_comment_kvs, cfg) -> str:
    ''' dir ex, show info'''
    if cls_name in already_pr:
        # print(f'### SKIP pred {cls_name}\n', file=sys.stderr)
        return ''
    else:
        already_pr.add(cls_name)
        # print(f'\n### SHOWINFO {cls_name}', file=sys.stderr)

    objxx = [repr(i) for i in dir(object())]
    xx = dir(ex)
    o_attrs = []
    o_methods = []
    o_unknowns = []
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
            elif str(typevv).startswith(cfg.win32comclsn):
                clsn = str(typevv).replace("<class '", '').replace("'>", '').split('.')[-1]
                all_cls[clsn] = vv
                o_attrs.append((i, typevv))
            else:
                o_unknowns.append((i, typevv))
        except AttributeError:
            e_noattr.append(i)
        except Exception as e:
            e_ee.append((i, e))

    cls_def_str = conv2cls(ex, cls_name, o_attrs, o_methods, o_unknowns, e_noattr, e_ee, cfg,
                           all_see_type, unparsed_type, obj2methods, clsn_member_comment_kvs)
    return cls_def_str


def output_all_cls_pyi(all_cls: dict[str, any], unparsed_type: dict, finded_type: dict, all_cls_def_strs: list, all_see_type):
    ''' all_cls.each -->   output  to pyi file'''
    ff = StringIO()

    # head
    print(f"# just stub file, import it , declare app obj, get ide auto type hit\n", file=ff)
    print(f"import os", file=ff)
    print(f"import sys", file=ff)
    print(f"import datetime", file=ff)
    print(file=ff)

    # output all unparsed
    print('# TODO FIXME unparsed', file=ff)
    for i in sorted(unparsed_type):
        print(f'class {i}: pass', file=ff)
    print(file=ff)

    # output all nofind define
    print('# TODO FIXME nofind define', file=ff)
    for i in sorted(all_see_type):
        if i not in finded_type:
            print(f'class {i}: pass', file=ff)
    print(file=ff)

    # print each class define
    for i in range(len(all_cls_def_strs)):
        data = all_cls_def_strs[i]  # order or reverse
        print(data, file=ff)

    print(f'## PRINT  {len(all_cls)}\n', file=sys.stderr)

    # get all
    data = ff.getvalue()

    # replace CoCobase
    hasbaseclss = list(filter(lambda x: True if x[0] == '_' else False, all_cls.keys()))
    for _base in hasbaseclss:
        cls = _base[1:]
        data = data.replace(f'class {cls}:', f'class {cls}({_base}):')

    # output all define
    print(data)

    print('all cls', len(all_cls), sorted(all_cls.keys()), file=sys.stderr)
    print('base cls', len(hasbaseclss), sorted(hasbaseclss), file=sys.stderr)


def output_all_cls_pyfake(fn: str, all_cls: dict[str, CDispatch]):
    ''' all_cls.each -> "class XX: pass"'''
    with open(fn, 'w') as ff:
        for cl in all_cls:
            print(f'class {cl}: pass', file=ff)
    print(f'output to py class info file: {fn}', file=sys.stderr)


# ------------------------------------------------------------------------


def runit(cfg: Cfg):
    obj2methods: dict[str, dict[str, str]] = {}
    obj2parameters: dict[str, dict[str, str]] = {}
    obj2unknow: dict[str, dict[str, str]] = {}

    all_see_type: dict[str, int] = {}
    finded_type: dict[str, int] = {}
    unparsed_type: dict[str, int] = {}
    clsn_member_comment_kvs: dict[str, str] = {}

    all_cls: dict[str, CDispatch] = {}
    already_pr: set[str] = set()

    all_cls_def_strs: list[str] = []

    load_apis(cfg.api_fn, obj2methods, obj2parameters, obj2unknow)
    load_comments(cfg.api_comment_fn, clsn_member_comment_kvs)
    print('------------------------------------------', file=sys.stderr)

    exapp = Dispatch('excel.application')
    exapp.Visible = True

    wb = exapp.Workbooks.Add()
    ws = exapp.ActiveSheet
    ws.Range("A1:B3").Value = [[2, 4], [5, 6], [4, 6]]
    ws.Shapes.AddChart()

    cls_name = 'Application'
    all_cls[cls_name] = exapp

    for i in range(10000):
        if all_cls.keys() == already_pr:
            print(f"## ALL DONE at {i}", file=sys.stderr)
            break
        for clsn in list(all_cls):
            exx = all_cls[clsn]
            cls_def_str = showinfo(clsn, exx, all_cls, already_pr, all_see_type,
                                   unparsed_type, obj2methods, clsn_member_comment_kvs, cfg)
            # print(f'for {i}   {clsn}  {len(cls_def_str)}', file=sys.stderr)
            if cls_def_str:
                finded_type[clsn] = 1
                all_cls_def_strs.append(cls_def_str)
    print('------------------------------------------', file=sys.stderr)

    output_all_cls_pyi(all_cls, unparsed_type, finded_type, all_cls_def_strs, all_see_type)
    output_all_cls_pyfake(cfg.out_pyfile_fn, all_cls)

    ws.Name = 'test'
    wb.Saved = True


if __name__ == '__main__':
    cfg = Cfg()
    check_or_exit(cfg)
    runit(cfg)
