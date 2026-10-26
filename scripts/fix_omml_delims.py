#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""fix_omml_delims — 修复 docx 里「右定界符缺失」的畸形公式组

症状：分段函数 / 不等式组左边是大括号 `{`，Word 里右边却多出一个撑大的 `)`。

根因见 check_omml_delims.py 的说明：OMML 里 m:begChr 默认 `(`、m:endChr 默认 `)`，
XSL 只在「不等于默认值」时才写出来。`\begin{cases}` 本该是
`begChr='{' + endChr=''`（空串＝不画右括号），那个 `endChr` 一旦被删掉，
Word 就套用默认的 `)`。

本脚本只补**能确定配对**的那几种左定界符，其余（如 `[-2,4)` 这类半开区间）不动：

    {  ->  ''      大括号不画右半
    |  ->  |       绝对值
    ‖  ->  ‖
    ⌈  ->  ⌉       ⌊ -> ⌋     ⟨ -> ⟩     ⌜ -> ⌟     ⟦ -> ⟧     〖 -> 〗

用法
----
    python fix_omml_delims.py <文件.docx>              # 另存 <名>_fixed.docx
    python fix_omml_delims.py <目录>                    # 目录下所有 docx 批量
    python fix_omml_delims.py <文件.docx> --inplace     # 原地改（先备份 .bak）
"""
import argparse
import os
import shutil
import sys
import zipfile

from lxml import etree

M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'

END_FOR = {
    '{': '', '|': '|', '‖': '‖',
    '⌈': '⌉', '⌊': '⌋', '⟨': '⟩',
    '⌜': '⌟', '⟦': '⟧', '〖': '〗',
}


def fix_root(root):
    """原地补 endChr，返回修复处数"""
    n = 0
    for d in root.iter(M + 'd'):
        dPr = d.find(M + 'dPr')
        if dPr is None:
            continue
        beg = dPr.find(M + 'begChr')
        end = dPr.find(M + 'endChr')
        if beg is None or end is not None:
            continue
        bv = beg.get(M + 'val')
        if bv not in END_FOR:
            continue                      # 不确定的（如 '['、'('）一律不碰

        val = END_FOR[bv]
        # 大括号要看分隔符：cases 是 sepChr=''（不画右括号），
        # 而 {1,2,3} 这种集合是 sepChr=','（该有右括号）
        if bv == '{':
            sep0 = dPr.find(M + 'sepChr')
            sv = sep0.get(M + 'val') if sep0 is not None else None
            val = '}' if sv == ',' else ''

        e = etree.SubElement(dPr, M + 'endChr')
        e.set(M + 'val', val)
        # schema 要求 dPr 内元素按 begChr, sepChr, endChr 顺序；挪到 sepChr 之后
        sep = dPr.find(M + 'sepChr')
        if sep is not None:
            dPr.remove(e)
            sep.addnext(e)
        n += 1
    return n


def rewrite_zip(src, dst, new_doc_xml):
    zin = zipfile.ZipFile(src)
    try:
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    data = new_doc_xml
                zout.writestr(item, data)
    finally:
        zin.close()


def fix_docx(path, inplace=False):
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read('word/document.xml')
    except (zipfile.BadZipFile, KeyError) as e:
        print('  跳过（不是正常的 docx）: %s' % e)
        return 0

    root = etree.fromstring(xml)
    n = fix_root(root)
    if n == 0:
        return 0

    new_xml = etree.tostring(root, xml_declaration=True,
                             encoding='UTF-8', standalone=True)
    if inplace:
        bak = path + '.bak'
        if not os.path.exists(bak):
            shutil.copy2(path, bak)
        tmp = path + '.tmp'
        rewrite_zip(path, tmp, new_xml)
        os.replace(tmp, path)
        print('  已修复 %d 处（原文件备份为 %s）' % (n, os.path.basename(bak)))
    else:
        base, ext = os.path.splitext(path)
        dst = base + '_fixed' + ext
        rewrite_zip(path, dst, new_xml)
        print('  修复 %d 处 -> %s' % (n, os.path.basename(dst)))
    return n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('target', help='docx 文件或目录')
    ap.add_argument('--inplace', action='store_true', help='原地修改（自动备份 .bak）')
    a = ap.parse_args()

    if os.path.isdir(a.target):
        files = [os.path.join(dp, f)
                 for dp, _, fn in os.walk(a.target)
                 for f in fn
                 if f.lower().endswith('.docx') and not f.startswith('~$')]
    else:
        files = [a.target]

    total_files = total_fixes = 0
    for p in files:
        n = fix_docx(p, a.inplace)
        if n:
            total_files += 1
            total_fixes += n
            print('  ↑ %s' % p)

    if total_files == 0:
        print('没有需要修的文件（全部正常）。')
        return 0
    print('\n共修 %d 个文件、%d 处。' % (total_files, total_fixes))
    return 0


if __name__ == '__main__':
    sys.exit(main())
