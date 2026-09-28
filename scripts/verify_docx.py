# -*- coding: utf-8 -*-
"""校验生成的 docx：公式对象数、是否残留降级文本、表格结构、LaTeX 残留、插图清单。"""
import re
import sys
import zipfile

M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

# 正文里不该出现的 LaTeX 命令——出现即说明公式被降级成了纯文本
_LATEX_CMD = re.compile(
    r'\\(frac|d?sqrt|left|right|sum|int|prod|lim|begin|end|infty|partial|nabla'
    r'|alpha|beta|gamma|delta|theta|lambda|mu|pi|sigma|phi|omega|Delta|Omega|Sigma|Pi'
    r'|times|cdot|div|pm|mp|leq|geq|neq|approx|equiv|sim|propto'
    r'|subset|subseteq|supset|supseteq|in|notin|cup|cap|varnothing|emptyset'
    r'|mathbb|mathcal|mathrm|text|vec|bar|hat|dot|tilde|overline|underline|quad)'
    r'[a-zA-Z]*'
)


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

    # LaTeX 命令残留（公式降级成纯文本的典型特征，如 \\frac 被当正文写出）
    latex_residue = []
    for t in texts:
        m = _LATEX_CMD.search(t)
        if m:
            latex_residue.append((m.group(0), t.strip()[:60]))

    # 空 oMath
    empty = 0
    for o in body.iter(M + 'oMath'):
        if not ''.join(o.itertext()).strip():
            empty += 1

    # 插图清单（与"应插图的题号"人工对账：数量、序号、尺寸、来源）
    imgs = []
    for i, dr in enumerate(body.iter(W + 'drawing'), 1):
        ext = dr.find('.//' + WP + 'extent')
        blip = dr.find('.//' + A + 'blip')
        rid = blip.get(R + 'embed') if blip is not None else '?'
        target = '?'
        try:
            rel = doc.part.rels[rid]
            target = rel.target_ref
        except KeyError:
            pass
        if ext is not None:
            cx, cy = int(ext.get('cx', 0)), int(ext.get('cy', 0))
            size = '%.1fx%.1fcm' % (cx / 360000, cy / 360000)
        else:
            size = '?'
        imgs.append((i, size, target.rsplit('/', 1)[-1]))

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
    print('LaTeX 残留    : %d' % len(latex_residue))
    print('插图          : %d 幅' % len(imgs))
    print('zip 完整性    : %s' % ('OK' if badzip is None else badzip))
    print('包内条目      : %d' % len(names))
    for cmd, t in latex_residue[:8]:
        print('   [LaTeX 残留] %s <- %s' % (cmd, t))
    for t in bad[:5]:
        print('   [降级] %s' % t)
    for t in junk[:5]:
        print('   [残留] %s' % t)
    for i, size, name in imgs:
        print('   [插图 %d] %s  %s' % (i, size, name))
    return 0 if (not bad and not junk and empty == 0 and not latex_residue) else 1


if __name__ == '__main__':
    sys.exit(main(sys.argv[1]))
