# -*- coding: utf-8 -*-
"""把 template.html、ExcelJS、lease_core.js 組成一個可離線使用的單檔 HTML。

    python build.py   →  ../縣有地租約連動工具.html
"""
import os

here = os.path.dirname(os.path.abspath(__file__))


def read(*p):
    with open(os.path.join(here, *p), encoding='utf-8') as f:
        return f.read()


html = read('template.html')
html = html.replace('/*EXCELJS*/', read('vendor', 'exceljs.min.js'), 1)
html = html.replace('/*CORE*/', read('lease_core.js'), 1)
out = os.path.join(here, '..', '縣有地租約連動工具.html')
with open(out, 'w', encoding='utf-8', newline='\n') as f:
    f.write(html)
print('built', os.path.abspath(out), f'{len(html.encode("utf-8")) / 1024:.0f} KB')
