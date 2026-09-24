#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
mtef2docx.py —— MathType 讲义 → Word 原生公式（端到端）

把 docx 里的 **MathType OLE 对象** 整体替换成 Word 原生 oMath 公式：

    word/embeddings/*.bin
        → OLE 复合文档 → Equation Native 流 → MTEF v5 二进制
        → (mtef.py)     完整对象树
        → (mtef.py)     LaTeX
        → (latex2omml)  OMML
        → 替换 <w:object>       Word 原生可编辑公式

替换后公式双击可编辑、能改字号、打印不发虚；原文件**不改动**。

用法
----
    # 1. 先体检，不动文件（默认只读）
    python scripts/mtef2docx.py "6.2.6n.docx"

    # 2. 看转换抽样
    python scripts/mtef2docx.py "6.2.6n.docx" --sample 15

    # 3. 确认后生成，输出到独立文件
    python scripts/mtef2docx.py "6.2.6n.docx" --apply

    # 4. 指定输出 / 批量
    python scripts/mtef2docx.py "6.2.6n.docx" --apply -o "新版.docx"
    python scripts/mtef2docx.py "C:\\讲义目录" --apply --outdir "C:\\输出"

    # 5. 导出公式清单
    python scripts/mtef2docx.py "6.2.6n.docx" --json eqs.json
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import mtef
from mtef import MTEFParser, mtef_from_ole, to_latex
from latex2omml import latex_to_omml

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn

# OLE 对象的 ProgID 前缀，用来判断「这是不是 MathType」
# 见过的有：Equation.DSMT4（MathType 6/7）、Equation.3（公式编辑器 3.0）、
#           Equation.KSEE3（MathType for Mac 老版本）
MATHTYPE_PROGIDS = (
    'equation.dsmt4',
    'equation.3',
    'equation.ksee3',
    'mathtype',
)

R_ID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

# python-docx 的 nsmap 里没有 'o' 前缀（它只注册了 w/r/m 等常用几个），
# 所以 OLEObject 必须用完整命名空间 URI 来找，qn('o:OLEObject') 会 KeyError
NS_O = 'urn:schemas-microsoft-com:office:office'
Q_OLEOBJECT = '{%s}OLEObject' % NS_O


# ---------------------------------------------------------------- 核心

def collect_objects(doc):
    """收集文档里所有 MathType OLE 对象

    返回 [{'elem': <w:object>, 'parent': 父元素, 'part': 'oleObjectN.bin',
           'rid': 'rIdN'}, ...]
    """
    items = []
    body = doc.element.body
    for obj in body.iter(qn('w:object')):
        ole = obj.find(Q_OLEOBJECT)
        if ole is None:
            continue
        progid = (ole.get('ProgID') or '').strip().lower()
        if not any(p in progid for p in MATHTYPE_PROGIDS):
            continue                       # 不是 MathType（可能是 Excel 图表等）
        rid = ole.get(R_ID)
        if not rid:
            continue
        rel = doc.part.rels.get(rid)
        if rel is None or rel.is_external:
            continue
        try:
            part = rel.target_part
            blob = part.blob
        except Exception:
            continue
        partname = str(part.partname)
        if '/embeddings/' not in partname:
            continue
        items.append({
            'elem': obj,
            'parent': obj.getparent(),
            'part': partname.rsplit('/', 1)[-1],
            'rid': rid,
            'blob': blob,
        })
    return items


def parse_ole(item):
    """从一个 embedding 的字节解出 (结构翻译的 LaTeX, 原始 TeX 源码)

    tex_input 是 MathType 6+ 存在 FUTURE 0x66 里的**作者原始 TeX 输入**，
    没有经过结构翻译的启发式猜测，通常更准（例如 {{A},{B}} 不会被压平成 A,B）。
    不是每个公式都有（手写方式建的就没有），没有时是 None。
    """
    p = MTEFParser(mtef_from_ole(item['blob']))
    node = p.parse()
    return to_latex(node), p.tex_input


def try_omml(latex):
    """LaTeX → OMML 的 OMML 字节串；转换失败返回 None（不抛异常）"""
    try:
        return latex_to_omml(latex).decode('utf-8')
    except Exception:
        return None


def pick_latex(latex, tex_input):
    """在「原始 TeX」和「结构翻译」之间选一个能用的

    优先 tex_input（更准），但前提是它能成功转成 OMML ——
    MathType 的 TeX Input Language 不是 100% 标准 LaTeX，
    个别语法 latex2mathml 不认，这时回退到结构翻译。
    """
    if tex_input:
        xml = try_omml(tex_input)
        if xml:
            return tex_input, xml
    return latex, try_omml(latex)


def build(doc, items, on_error='keep'):
    """把 OLE 对象逐个换成原生公式

    on_error: 'keep'  失败时保留原对象（推荐）
              'text'  失败时降级为 [LaTeX 文本]
    返回 {'ok': [...], 'fail': [...], 'skip': [...]}
    """
    from docx.oxml import parse_xml as _parse_xml

    result = {'ok': [], 'fail': [], 'skip': []}
    for it in items:
        obj, parent = it['elem'], it['parent']
        try:
            latex, tex = parse_ole(it)
        except Exception as e:
            # 失败一律保留原对象 —— 宁可留着 MathType 图，也不能把公式弄丢
            result['fail'].append({**it, 'reason': 'MTEF 解析失败: %s' % e})
            continue

        if not (latex or '').strip():
            result['skip'].append({**it, 'reason': '公式为空'})
            continue

        used, omml_xml = pick_latex(latex, tex)
        if not omml_xml:
            result['fail'].append(
                {**it, 'reason': 'LaTeX→OMML 失败: %s' % used})
            continue

        try:
            new_elem = _parse_xml(omml_xml)
        except Exception as e:
            result['fail'].append({**it, 'reason': 'OMML 不是合法 XML: %s' % e})
            continue

        parent.replace(obj, new_elem)
        result['ok'].append({**it, 'latex': used, 'from_tex': used == tex
                             and tex is not None})

    return result


# ---------------------------------------------------------------- CLI

def scan(path, args):
    """只读体检"""
    try:
        doc = Document(path)
    except Exception as e:
        return {'file': path, 'error': '打开失败: %s' % e}

    items = collect_objects(doc)
    rec = {'file': path, 'objects': len(items), 'ok': 0, 'fail': 0,
           'skip': 0, 'from_tex': 0, 'samples': [], 'errors': []}

    for it in items:
        try:
            latex, tex = parse_ole(it)
        except Exception as e:
            rec['fail'] += 1
            rec['errors'].append('%s: %s' % (it['part'], e))
            continue
        used, xml = pick_latex(latex, tex)
        if not xml:
            rec['fail'] += 1
            rec['errors'].append('%s: 转 OMML 失败 %r' % (it['part'], used))
            continue
        rec['ok'] += 1
        if tex is not None and used == tex:
            rec['from_tex'] += 1
        if len(rec['samples']) < (args.sample or 8):
            rec['samples'].append((it['part'], used,
                                   'TeX' if used == tex else '结构'))

    return rec


def main():
    ap = argparse.ArgumentParser(
        description='MathType 讲义 → Word 原生公式（MTEF 解析，免 OCR）')
    ap.add_argument('target', help='.docx 文件或目录')
    ap.add_argument('--apply', action='store_true',
                    help='真正生成新 docx（默认只读体检）')
    ap.add_argument('-o', '--out', help='输出文件（单文件模式）')
    ap.add_argument('--outdir', help='输出目录（目录模式）')
    ap.add_argument('--sample', type=int, default=8, help='抽样条数，默认 8')
    ap.add_argument('--json', help='导出公式清单 JSON')
    args = ap.parse_args()

    if os.path.isdir(args.target):
        # 必须递归：讲义库通常是「导学案/第六章/...」这种多层结构，
        # 只扫顶层会漏掉绝大部分文档（实测顶层只有 7 份，实际有 214 份）
        files = []
        for dp, dn, fn in os.walk(args.target):
            dn[:] = [d for d in dn if not d.startswith('.')]
            for n in sorted(fn):
                if n.lower().endswith('.docx') and not n.startswith('~$'):
                    files.append(os.path.join(dp, n))
        files.sort()
    else:
        files = [args.target]

    if not files:
        print('没找到 .docx')
        sys.exit(1)

    total = {'objects': 0, 'ok': 0, 'fail': 0, 'skip': 0}
    all_eq = []

    for path in files:
        if not args.apply:
            rec = scan(path, args)
            if 'error' in rec:
                print('\n%s\n  [!] %s' % (os.path.basename(path), rec['error']))
                continue
            print('\n=== %s ===' % os.path.basename(path))
            print('  MathType 对象: %d  可转: %d（其中用原TeX %d）  失败: %d'
                  % (rec['objects'], rec['ok'], rec['from_tex'], rec['fail']))
            for e in rec['errors'][:10]:
                print('    [X] %s' % e)
            if rec['samples']:
                print('  --- 抽样 ---')
                for part, latex, src in rec['samples']:
                    print('    %-20s [%-2s] %s' % (part, src, latex))
            total['objects'] += rec['objects']
            total['ok'] += rec['ok']
            total['fail'] += rec['fail']
            continue

        # ---- --apply ----
        doc = Document(path)
        items = collect_objects(doc)
        res = build(doc, items)

        if args.out and len(files) == 1:
            out = args.out
        elif args.outdir:
            os.makedirs(args.outdir, exist_ok=True)
            out = os.path.join(args.outdir, os.path.basename(path))
        else:
            base, ext = os.path.splitext(path)
            out = base + '_原生公式' + ext

        doc.save(out)
        n_tex = sum(1 for r in res['ok'] if r.get('from_tex'))
        print('\n=== %s ===' % os.path.basename(path))
        print('  对象 %d → 成功 %d（用原TeX %d），失败 %d，跳过 %d'
              % (len(items), len(res['ok']), n_tex,
                 len(res['fail']), len(res['skip'])))
        for f in res['fail'][:10]:
            print('    [X] %s: %s' % (f['part'], f['reason']))
        print('  已保存 → %s' % out)

        total['objects'] += len(items)
        total['ok'] += len(res['ok'])
        total['fail'] += len(res['fail'])
        total['skip'] += len(res['skip'])
        all_eq.extend(res['ok'])

    print('\n' + '=' * 52)
    if args.apply:
        print('合计：对象 %d，成功 %d，失败 %d，跳过 %d'
              % (total['objects'], total['ok'], total['fail'], total['skip']))
        if not args.apply:
            print('\n（只读体检模式，加 --apply 才生成文件）')
    else:
        print('合计：对象 %d，可转 %d，失败 %d'
              % (total['objects'], total['ok'], total['fail']))
        print('\n这是只读体检，加 --apply 才生成新 docx（原文件不动）')

    if args.json:
        import json
        with open(args.json, 'w', encoding='utf-8') as f:
            json.dump(all_eq, f, ensure_ascii=False, indent=1)
        print('已导出 %s' % args.json)


if __name__ == '__main__':
    main()
