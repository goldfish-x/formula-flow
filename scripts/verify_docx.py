# -*- coding: utf-8 -*-
"""校验生成的 docx：公式对象数、是否残留降级文本、表格结构。"""
import re
import sys
import zipfile

M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def main(path):
    from docx import Document
    doc = Document(path)
    body = doc.element.body
    n_math = sum(1 for _ in body.iter(M + 'oMath'))
    n_mpara = sum(1 for _ in body.iter(M + 'oMathPara'))
    n_tbl = len(doc.tables)

    texts = []
    for p in doc.paragraphs:
        t = p.text
        texts.append(t)

    # 降级文本（公式转换失败时留下的 [...] 标记）
    bad = [t for t in texts if re.search(r'\[[^\]]{2,}\]', t)]

    # 定界符拼接件残留
    junk = [t for t in texts if re.search(r'[⎛⎜⎝⎡⎢⎣⎧⎨⎩⎪⎫⎬⎭⎮⎞⎟⎠⎤⎥⎦]', t)]

    # 空 oMath
    empty = 0
    for o in body.iter(M + 'oMath'):
        if not ''.join(o.itertext()).strip():
            empty += 1

    # docx 包完整性
    with zipfile.ZipFile(path) as z:
        badzip = z.testzip()
        names = z.namelist()

    print('文件          : %s' % path)
    print('段落          : %d（其中空段落 %d）'
          % (len(texts), sum(1 for t in texts if not t.strip())))
    print('表格          : %d' % n_tbl)
    print('原生公式      : 行内 %d，独立 %d' % (n_math, n_mpara))
    print('空公式对象    : %d' % empty)
    print('降级文本残留  : %d' % len(bad))
    print('定界符件残留  : %d' % len(junk))
    print('zip 完整性    : %s' % ('OK' if badzip is None else badzip))
    print('包内条目      : %d' % len(names))
    for t in bad[:5]:
        print('   [降级] %s' % t)
    for t in junk[:5]:
        print('   [残留] %s' % t)
    return 0 if (not bad and not junk and empty == 0) else 1


if __name__ == '__main__':
    sys.exit(main(sys.argv[1]))
