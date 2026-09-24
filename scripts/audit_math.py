# -*- coding: utf-8 -*-
"""
讲义公式体检：扫描 docx，统计「真公式」与「假公式」

真公式 = Word 原生 oMath 对象（可双击编辑）
假公式 = 图片里的公式 / 纯文本上标(^) / 纯文本分数(/) / 裸 Unicode 符号
"""
import os
import re
import sys
import zipfile
import glob

M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# 数学符号 Unicode 区段
MATH_RANGES = [
    (0x2200, 0x22FF),  # 数学运算符
    (0x2190, 0x21FF),  # 箭头
    (0x2A00, 0x2AFF),  # 补充数学运算符
    (0x0391, 0x03C9),  # 希腊字母
]
MATH_EXTRA = set('∅∈∉⊆⊇⊂⊃⫋⫌∩∪≠≤≥±×÷⋅√∫∑∏∞αβγδεζηθικλμνξπρστυφχψω')


def is_math_char(ch):
    o = ord(ch)
    if ch in MATH_EXTRA:
        return True
    for lo, hi in MATH_RANGES:
        if lo <= o <= hi:
            return True
    return False


def audit(path):
    try:
        z = zipfile.ZipFile(path)
    except Exception as e:
        return {'error': str(e)}

    try:
        xml = z.read('word/document.xml').decode('utf-8', 'ignore')
    except KeyError:
        return {'error': '不是有效的 docx'}

    # 1. 原生公式对象
    real = len(re.findall(r'<m:oMath[ >]', xml))

    # 2. 图片（公式可能被存成图）
    images = len(re.findall(r'<a:blip[ >]', xml)) + \
        len(re.findall(r'<v:imagedata[ >]', xml))
    # 图元文件（EMF/WMF）通常就是公式
    metafiles = len([n for n in z.namelist()
                     if n.startswith('word/media/')
                     and n.lower().endswith(('.emf', '.wmf'))])

    # 3. 纯文本内容
    texts = re.findall(r'<w:t[^>]*>(.*?)</w:t>', xml, re.S)
    plain = ''.join(texts)
    plain = (plain.replace('&amp;', '&').replace('&lt;', '<')
             .replace('&gt;', '>').replace('&quot;', '"'))

    # 4. 假公式特征
    caret_sup = len(re.findall(r'[A-Za-z0-9)]\^ ?\{?[A-Za-z0-9]', plain))  # x^2
    text_frac = len(re.findall(r'[A-Za-z0-9)]/[A-Za-z0-9(]', plain))       # a/b
    sqrt_text = plain.count('sqrt') + plain.count('√')
    bare_sym = sum(1 for ch in plain if is_math_char(ch))

    return {
        'real': real,
        'images': images,
        'metafiles': metafiles,
        'caret_sup': caret_sup,
        'text_frac': text_frac,
        'sqrt_text': sqrt_text,
        'bare_sym': bare_sym,
        'chars': len(plain),
    }


def main():
    if len(sys.argv) > 1:
        target = sys.argv[1]
    else:
        print('用法: python audit_math.py <讲义目录>')
        sys.exit(1)

    files = sorted(glob.glob(os.path.join(target, '*.docx')))
    print(f'扫描目录: {target}')
    print(f'文件数: {len(files)}\n')
    print(f'{"文件名":<34} {"真公式":>6} {"图片":>5} {"图元":>5} '
          f'{"^上标":>6} {"/分数":>6} {"裸符号":>6}')
    print('-' * 80)

    totals = dict(real=0, images=0, metafiles=0, caret_sup=0,
                  text_frac=0, sqrt_text=0, bare_sym=0)
    problems = []

    for f in files:
        r = audit(f)
        name = os.path.basename(f)
        if 'error' in r:
            print(f'{name:<34}  错误: {r["error"]}')
            continue
        for k in totals:
            totals[k] += r[k]
        flag = ''
        if r['real'] == 0 and (r['images'] or r['metafiles'] or r['bare_sym']):
            flag = '  <-- 无原生公式'
            problems.append((name, r))
        print(f'{name:<34} {r["real"]:>6} {r["images"]:>5} {r["metafiles"]:>5} '
              f'{r["caret_sup"]:>6} {r["text_frac"]:>6} {r["bare_sym"]:>6}{flag}')

    print('-' * 80)
    print(f'{"合计":<34} {totals["real"]:>6} {totals["images"]:>5} '
          f'{totals["metafiles"]:>5} {totals["caret_sup"]:>6} '
          f'{totals["text_frac"]:>6} {totals["bare_sym"]:>6}')

    print(f'\n=== 结论 ===')
    print(f'原生可编辑公式总计: {totals["real"]} 个')
    print(f'嵌入图片: {totals["images"]} 个，其中 EMF/WMF 图元: '
          f'{totals["metafiles"]} 个（图元通常就是公式）')
    print(f'纯文本上标(如 x^2): {totals["caret_sup"]} 处')
    print(f'纯文本分数(如 a/b): {totals["text_frac"]} 处')
    print(f'裸 Unicode 符号(文本层，非公式对象): {totals["bare_sym"]} 个')
    if problems:
        print(f'\n需要关注的文件（0 个原生公式但有数学内容）: {len(problems)} 个')
        for name, r in problems:
            print(f'  - {name}')


if __name__ == '__main__':
    main()
