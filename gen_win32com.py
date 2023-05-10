# coding: utf-8
import io
import json
import os.path
import sys


class Cfg():
    '''Config TODO  sync rb/py file'''
    OUT_EXCEL_INFO_FN = "excel.info.json"

    N_SUMM = "summary"
    N_meta = 'meta'

    N_co = "collections"
    N_co_doc = "co_doc"

    N_e = "enumerations"
    N_e_unique = "uniqued"
    N_e_doc = "e_doc"
    N_e_table_head = "e_table_head"
    N_e_table_rows = "e_table_rows"
    N_e_remarks_doc = "e_remarks_doc"
    N_e_table_format_incorrect = "e_ERROR"

    N_c = "classes"
    N_c_doc = "c_doc"
    N_c_remarks_doc = "c_remarks_doc"
    N_c_example_doc = "c_examples_doc"

    N_m = "methods"
    N_m_doc = "m_doc"
    N_m_return = "m_return"
    N_m_return_doc = "m_return_doc"
    N_m_parameters_doc = "m_parameters_doc"
    N_m_remarks_doc = "m_remarks_doc"
    N_m_example_doc = "m_example_doc"

    N_p = "properties"
    N_p_doc = "p_doc"
    N_p_type = "p_type"
    N_p_syntax_doc = "p_syntax_doc"
    N_p_return_doc = "p_return_doc"
    N_p_property_value_doc = "p_property_value_doc"
    N_p_remarks_doc = "p_remarks_doc"
    N_p_example_doc = "p_example_doc"

    OUT_EXCEL_PYI_FN = "oletype/excel.pyi"
    OUT_EXCEL_PY_FN = "oletype/excel.py"
    MYPKGNAME = "excel."
    MYPKGNAME = ""

    # python \Python311\Lib\site-packages\win32com\client\makepy.py
    #  Generating to C:\Users\user1\AppData\Local\Temp\gen_py\3.11\00020813-0000-0000-C000-000000000046x0x1x9.py
    # ole_class_pre: str = "<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9."
    # ole_class_pre2: str = "win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9."
    OLE_CLASS_PRE: str = "<class 'win32com."

    BUILTINS_CLASS = [clsn for clsn in dir(__builtins__) if type(getattr(__builtins__, clsn)) == type]
    BUILTINS_CLASS.append('None')
    OBJECT_PROPERTIES = [repr(i) for i in dir(object())]
    OLEOBJECT_SKIP_PROPERTY = ['CLSID', '__weakref__', '_oleobj_', 'coclass_clsid']
    OLEOBJECT_SKIP_EXCEPTION_PROPERTY = ['Visible']

    OUTFILE_HEADER = '''# coding: utf-8

# Generated by oletype, win32com(excel) object py/pyi file
#   for ide tips
#
# Usage:
#   from win32com.client import Dispatch
#   from oletype import excel
#   exapp = excel.Application = Dispatch('excel.application')  #declare app obj
#   exapp.   #  get ide auto type hit
'''
    IMPORTS = '''
from enum import IntFlag, unique
import datetime
'''
    num = 0


def check_or_exit():
    if not os.path.exists(Cfg.OUT_EXCEL_INFO_FN):
        print(f'not exists info.json file {Cfg.OUT_EXCEL_INFO_FN}')
        sys.exit(2)


def load_info_from_json(fn: str, kvs: dict):
    with open(fn, encoding='utf-8', errors='ignore') as ff:
        dd = ff.read()
        kvs.update(json.loads(dd))

    fs = f"from {fn} read {len(dd)} bytes"
    summ = f"{kvs[Cfg.N_SUMM]}, "
    collections = f"{len(kvs[Cfg.N_co])} collections"
    enumerations = f"{len(kvs[Cfg.N_e])} enumerations"
    classes = f"{len(kvs[Cfg.N_c])} classes"
    methods = f"{len(kvs[Cfg.N_c + '_' + Cfg.N_m])} methods"
    properties = f"{len(kvs[Cfg.N_c + '_' + Cfg.N_p])} properties"

    print(f"{fs}\n SUMMARY:  {summ}\n JSON loads:  {collections}, {enumerations}, {classes}, {methods}, {properties}.", file=sys.stderr)


def get_cln_class_comment(objinfo: dict) -> str | None:
    if objinfo:
        c_doc = objinfo.get(Cfg.N_c_doc)
        c_remarks_doc = objinfo.get(Cfg.N_c_remarks_doc)
        c_example_doc = objinfo.get(Cfg.N_c_example_doc)
        sss = []
        if c_doc:
            sss.append(c_doc)
        if c_remarks_doc:
            sss.append("#REMARKS:\n\n" + c_remarks_doc)
        if c_example_doc:
            sss.append("#EXAMPLE:\n\n" + c_example_doc)
        sss = '\n\n'.join(sss)
        if sss and sss[-1] in "'\"":
            sss += ' '
        return sss
    return None


def get_cln_method_comment(objinfo: dict) -> str | None:
    if objinfo:
        m_doc = objinfo.get(Cfg.N_m_doc)
        m_return_doc = objinfo.get(Cfg.N_m_return_doc)
        m_parameters_doc = objinfo.get(Cfg.N_m_parameters_doc)
        m_remarks_doc = objinfo.get(Cfg.N_m_remarks_doc)
        m_example_doc = objinfo.get(Cfg.N_m_example_doc)
        sss = []
        if m_doc:
            sss.append(m_doc)
        if m_parameters_doc:
            sss.append('#PARAMETERS:\n\n' + m_parameters_doc)
        if m_return_doc:
            sss.append('#RETURN-VALUE: ' + m_return_doc)
        if m_remarks_doc:
            sss.append('#REMARKS:\n\n' + m_remarks_doc)
        if m_example_doc:
            sss.append('#EXAMPLE:\n\n' + m_example_doc)
        sss = '\n\n'.join(sss)
        if sss and sss[-1] in "'\"":
            sss += ' '
        return sss
    return None


def get_cln_property_comment(objinfo: dict) -> str | None:
    if objinfo:
        p_doc = objinfo.get(Cfg.N_p_doc)
        p_syntax_doc = objinfo.get(Cfg.N_p_syntax_doc)
        p_return_doc = objinfo.get(Cfg.N_p_return_doc)
        p_property_value_doc = objinfo.get(Cfg.N_p_property_value_doc)
        p_remarks_doc = objinfo.get(Cfg.N_p_remarks_doc)
        p_example_doc = objinfo.get(Cfg.N_p_example_doc)
        sss = []
        if p_doc:
            sss.append(p_doc)
        if p_syntax_doc:
            sss.append('#SYNTAX:\n\n' + p_syntax_doc)
        if p_return_doc:
            sss.append('#RETRUN-VALUE: ' + p_return_doc)
        if p_property_value_doc:
            sss.append('#PROPERTY-VALUE: ' + p_property_value_doc)
        if p_remarks_doc:
            sss.append('#REMARKS:\n\n' + p_remarks_doc)
        if p_example_doc:
            sss.append('#EXAMPLE:\n\n' + p_example_doc)
        sss = '\n\n'.join(sss)
        if sss and sss[-1] in "'\"":
            sss += ' '
        return sss
    return None


def conv2cls(cls_name: str,
             ole_info_kvs: dict,
             unfoundcls: set,
             ) -> str:
    '''from dict find cls attrs, methods, unknown properties, and errors'''
    cl_info_kvs = ole_info_kvs[Cfg.N_c][cls_name]
    Cfg.num += 1

    ff = io.StringIO()

    p_ns = []
    m_ns = []
    for na in sorted(cl_info_kvs.keys()):
        if na == Cfg.N_meta:
            continue
        objinfo = cl_info_kvs[na]
        if dict != type(objinfo):
            continue
        if Cfg.N_p_doc in objinfo:
            p_ns.append(na)
        elif Cfg.N_m_doc in objinfo:
            m_ns.append(na)
        else:
            print(f'unknown objinfo {objinfo.keys()}', file=sys.stderr)

    print(f" {cls_name} \t {len(p_ns) } {len(m_ns)}", file=sys.stderr)
    if len(p_ns) > 0:
        print('  def __init__(self):', file=ff)
    for na in p_ns:
        objinfo = cl_info_kvs[na]
        ty = objinfo[Cfg.N_p_type]
        if ty:
            for ii in ty.split('|'):
                ii = ii.strip()
                unfoundcls.add(ii)
        comments = get_cln_property_comment(objinfo)
        ind = '    '
        indsuff = f"\n{ind}" if (comments and ('\n' in comments or comments[-1] in "'\"")) else ''
        comments = f"\n{ind}'''{comments}{indsuff}'''\n" if comments else ''
        print(f"    self.{na}: {ty}{comments}", file=ff)
    print(file=ff)

    print(f'# {cls_name} ns', file=ff)
    for na in m_ns:
        objinfo = cl_info_kvs[na]
        ret = objinfo.get(Cfg.N_m_return)
        ars = ''  # TODO FIXME
        if ret is not None:
            if ret:
                for ii in str(ret).split('|'):
                    ii = ii.strip()
                    unfoundcls.add(ii)
            rets = ' -> ' + str(ret)
        else:
            rets = ''
        comments = get_cln_method_comment(objinfo)
        ind = '    '
        indsuff = f"\n{ind}" if (comments and ('\n' in comments or comments[-1] in "'\"")) else ''
        comments = f"\n{ind}'''{comments}{indsuff}'''\n" if comments else '  pass\n'
        print(f"  def {na}(self{(', ' + ars) if ars else ''}){rets}:{comments}", file=ff)
    print(f'# {cls_name} ns end', file=ff)
    print(file=ff)

    var_methodss = ff.getvalue()

    # header
    ff = io.StringIO()
    print(file=ff)
    # print(f'# num={Cfg.num}', file=ff)
    print(f'class {cls_name}:', file=ff)
    comments = get_cln_class_comment(cl_info_kvs)
    if comments:
        ind = '  '
        indsuff = f"\n{ind}" if ('\n' in comments or comments[-1] in "'\"") else ''
        print(f"{ind}'''{comments}{indsuff}'''\n", file=ff)
    print(file=ff)


    # var methods
    print(var_methodss, file=ff)
    print(file=ff)

    # print(f'# Summary "{ex.__class__.__mro__}", attrs:{len(o_attrs)}, methods:{len(o_methods)}, unknowns:{len(o_unknowns)},   ok:{len(o_attrs) + len(o_methods) + len(o_unknowns)}, e_noattr:{len(e_noattr)}, e_eerror:{ len(e_ee)}', file=ff)
    return ff.getvalue()


def out_collection(ff, ole_info_kvs: dict):
    for k in ole_info_kvs[Cfg.N_co]:
        kvs: dict = ole_info_kvs[Cfg.N_co][k]
        print(f'class {k}:', file=ff)
        e_doc = kvs.get(Cfg.N_co_doc)
        sss = []
        if e_doc:
            sss.append(e_doc)
        comments = "\n\n".join(sss)
        ind = '  '
        indsuff = f'\n{ind}' if ('\n' in comments or comments[-1] in "'\"") else ''
        print(f"{ind}'''{comments}{indsuff}'''\n", file=ff)
    pass


def out_enumeration(ff, ole_info_kvs: dict):
    for k in ole_info_kvs[Cfg.N_e]:

        kvs: dict = ole_info_kvs[Cfg.N_e][k]

        if kvs.get(Cfg.N_e_unique):
            print(f'@unique', file=ff)
        print(f'class {k}(IntFlag):', file=ff)

        e_doc = kvs.get(Cfg.N_e_doc)
        e_remarks_doc = kvs.get(Cfg.N_e_remarks_doc)
        sss = []
        if e_doc:
            sss.append(e_doc)
        if e_remarks_doc:
            sss.append('#REMARKS:\n\n' + e_remarks_doc)
        comments = "\n\n".join(sss)
        ind = '  '
        indsuff = f'\n{ind}' if ('\n' in comments or comments[-1] in "'\"") else ''
        print(f"{ind}'''{comments}{indsuff}'''\n", file=ff)

        e_table_rows = kvs.get(Cfg.N_e_table_rows)
        for ii in e_table_rows:
            if not ii:
                continue
            try:
                n, v, desc = ii[0], ii[1], ii[2]
                print(f"  {n} = {v}", file=ff)
                print(
                    f"  '''{desc}{ '  '+' '.join([str(i) for i in ii[3:]]) if len(ii)>3 else '' }'''", file=ff)
            except:
                print(f'ERR ii {ii}', file=sys.stderr)
        print(file=ff)
    pass


def out_collection_enumeration(ff, ole_info_kvs: dict) -> list:
    # out collection
    coks = ole_info_kvs[Cfg.N_co]
    print(f'# list collection  {len(coks)}', file=ff)
    out_collection(ff, ole_info_kvs)
    print(f'# list collection  end', file=ff)
    print(file=ff)

    # out enumeration
    eks = ole_info_kvs[Cfg.N_e]
    print(f'# list enumeration  {len(eks)}', file=ff)
    out_enumeration(ff, ole_info_kvs)
    print(f'# list enumeration  end', file=ff)
    print(file=ff)

    c = []
    c.extend(coks.keys())
    c.extend(eks.keys())
    return c


def output_to_pyi_typehints(fn: str,
                            ole_info_kvs: dict,
                            ):
    '''all_cls_kvs.each -->   output  to pyi file'''
    ff = io.StringIO()

    # head
    print(Cfg.OUTFILE_HEADER, file=ff)
    print(file=ff)
    print(Cfg.IMPORTS, file=ff)
    print(file=ff)

    coeks = out_collection_enumeration(ff, ole_info_kvs)
    print(file=ff)

    unfoundcls = set()

    clsns = sorted(ole_info_kvs[Cfg.N_c].keys())
    print(f'# ole cls  {len(clsns)}', file=ff)
    for clsn in clsns:
        clss = conv2cls(clsn, ole_info_kvs, unfoundcls)
        print(clss, file=ff)
    print(f'# ole cls end', file=ff)
    print(file=ff)


    # unfoundcls
    nns = sorted([str(i) for i in unfoundcls])
    print(f'# unfoundcls', file=ff)
    for cls_name in nns:
        if not cls_name:
            continue
        if '.' not in cls_name:
            if cls_name in Cfg.BUILTINS_CLASS or cls_name in clsns or cls_name in coeks:
                continue
            print(f'class {cls_name}: pass', file=ff)
        else:
            if cls_name[0] in 'abcdefghijklmnopqrstuvwxyz':
                continue
            nss = cls_name.split('.')
            indi = 0
            for i in range(len(nss)-1):
                cls_name = nss[i]
                if cls_name in Cfg.BUILTINS_CLASS:
                    continue
                print(f'{" "*indi}class {cls_name}:', file=ff)
                indi += 2
            print(f'{" "*indi}class {nss[-1]}: pass', file=ff)
    print(f'# unfoundcls  end', file=ff)
    print(file=ff)

    # get all
    data = ff.getvalue()

    # output all define
    with open(fn, 'w', encoding='utf-8') as fout:
        print(data, file=fout)


def output_to_py_src(fn: str,
                     ole_info_kvs: dict
                     ):
    '''all_cls_kvs.each -> "class XX: pass"'''
    with open(fn, 'w', encoding='utf-8') as ff:
        print(Cfg.OUTFILE_HEADER, file=ff)
        print(file=ff)
        print(Cfg.IMPORTS, file=ff)

        out_collection_enumeration(ff, ole_info_kvs)
        print(file=ff)

        clsns = sorted(ole_info_kvs[Cfg.N_c].keys())
        for clsn in clsns:
            print(f'class {clsn}: pass', file=ff)
        print(file=ff)


# ------------------------------------------------------------------------


def runit():
    ole_info_kvs: dict[str, str] = {}

    load_info_from_json(Cfg.OUT_EXCEL_INFO_FN, ole_info_kvs)
    print('------------------------------------------', file=sys.stderr)

    output_to_pyi_typehints(Cfg.OUT_EXCEL_PYI_FN, ole_info_kvs)
    output_to_py_src(Cfg.OUT_EXCEL_PY_FN, ole_info_kvs)


if __name__ == '__main__':
    check_or_exit()
    runit()
