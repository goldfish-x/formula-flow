# -*- coding: utf-8 -*-
"""
mtef_v3_probe — 探查 docx 里的 MathType 对象，给「MTEF v3 无法解析」一条可走的路

背景
----
mtef2docx.py 只能解析 MathType **v5**。Word 里那种很老的
「公式编辑器 3.0」对象（ProgID = Equation.KSEE3 / Equation.3）是 **MTEF v3**，
解析器会明确报错并让人去 Word 里转换——但如果手上没有装 Office，
或者不想逐个双击转换，就得另找办法。

本工具不试图解析 v3，它把判定公式所需的材料一次性摆出来：
  1. 每个 OLE 对象的 MTEF 版本、是否含原始 TeX（有些对象存了 TeX）
  2. 配套的 WMF/EMF 预览图 → 渲染成 PNG（GDI+ / LibreOffice / Inkscape）
  3. 从图元里精确提取字符（不靠 OCR，符号不会看错）

然后人（或多模态模型）看 PNG 就能把公式写回 LaTeX，再用 latex2omml 生成原生公式。

用法
----
    python mtef_v3_probe.py 文档.docx [-o 输出目录]

输出
----
    <输出目录>/ole_N.bin          原始对象
    <输出目录>/formula_N.png      渲染图（用来读公式）
    <输出目录>/probe.json         版本 / 符号 / 图片路径清单
"""
import argparse
import json
import os
import re
import sys
import zipfile

SKILL_SCRIPTS = os.path.dirname(os.path.abspath(__file__))
if SKILL_SCRIPTS not in sys.path:
    sys.path.insert(0, SKILL_SCRIPTS)

import emf2png          # noqa: E402
import wmf_symbols      # noqa: E402

V3_PROGIDS = ('Equation.3', 'Equation.KSEE3', 'Equation.DSMT4')


def mtef_version(bin_path):
    """读 OLE 里 Equation Native 流的 MTEF 版本号。

    结构：EQNOLEFILEHDR(28 字节) 之后就是 MTEF 数据，
    首字节 = MTEF 版本（2/3/5）。
    """
    try:
        import olefile
    except ImportError:
        return None, 'olefile 未安装（pip install olefile）'
    if not olefile.isOleFile(bin_path):
        return None, '不是 OLE 复合文档'
    ole = olefile.OleFileIO(bin_path)
    if not ole.exists('Equation Native'):
        return None, '没有 Equation Native 流'
    data = ole.openstream('Equation Native').read()
    ole.close()
    if len(data) < 32:
        return None, 'Equation Native 流过短'
    cb = int.from_bytes(data[0:4], 'little')
    ver = data[cb] if cb < len(data) else data[0]
    return ver, f'头部 {cb} 字节，MTEF v{ver}'


def symbols_of(path):
    """用图元解析取字符（不靠 OCR，符号不会看错）"""
    try:
        items = wmf_symbols.extract(path)
    except Exception:                       # noqa: BLE001
        return None, []
    if not items:
        return '', []
    text = ' '.join(it['text'] for it in items)
    try:
        flags = wmf_symbols.summarize(items)
    except Exception:                       # noqa: BLE001
        flags = []
    return text, flags


def main():
    ap = argparse.ArgumentParser(description='探查 docx 里的 MathType 对象')
    ap.add_argument('docx')
    ap.add_argument('-o', '--outdir', default=None)
    args = ap.parse_args()

    base = os.path.splitext(args.docx)[0]
    outdir = args.outdir or (base + '_mtef')
    os.makedirs(outdir, exist_ok=True)

    z = zipfile.ZipFile(args.docx)
    bins = sorted(n for n in z.namelist()
                  if n.startswith('word/embeddings/') and n.lower().endswith('.bin'))
    media = sorted(n for n in z.namelist()
                   if n.startswith('word/media/')
                   and n.lower().endswith(('.wmf', '.emf')))

    print(f'\n=== {os.path.basename(args.docx)} ===')
    print(f'嵌入对象 {len(bins)} 个，图元 {len(media)} 个')
    print(f'渲染后端: {emf2png.backend_name()}')

    report = {'docx': args.docx, 'objects': []}

    for i, name in enumerate(bins, 1):
        stem = f'ole_{i}'
        bin_path = os.path.join(outdir, stem + '.bin')
        with open(bin_path, 'wb') as f:
            f.write(z.read(name))
        ver, note = mtef_version(bin_path)

        # 同名序号的图元（Word 通常按顺序配对）
        png = None
        if i - 1 < len(media):
            m = media[i - 1]
            ext = os.path.splitext(m)[1]
            img_path = os.path.join(outdir, stem + ext)
            with open(img_path, 'wb') as f:
                f.write(z.read(m))
            png = os.path.join(outdir, stem + '.png')
            try:
                emf2png.convert(img_path, png, scale=3)
            except Exception as e:          # noqa: BLE001
                print(f'  渲染失败: {e}')
                png = None

        sym, flags = symbols_of(img_path) if i - 1 < len(media) else (None, [])

        print(f'\n[{i}] {os.path.basename(name)}')
        print(f'    版本   : {note}')
        print(f'    图元   : {os.path.basename(img_path) if i - 1 < len(media) else "无"}')
        if png:
            print(f'    渲染图 : {png}')
        if sym:
            print(f'    字符   : {sym}')
        if flags:
            print(f'    特征   : {"、".join(flags)}')

        report['objects'].append({
            'bin': bin_path, 'mtef_version': ver,
            'png': png, 'symbols': sym, 'flags': flags,
            'supported': ver == 5,
        })

    with open(os.path.join(outdir, 'probe.json'), 'w', encoding='utf-8') as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    print('\n' + '-' * 60)
    n5 = sum(1 for o in report['objects'] if o['supported'])
    print(f'可自动解析(v5) {n5} 个，需人工确认 {len(bins) - n5} 个')
    if len(bins) - n5:
        print('下一步：看上面的渲染图确认公式 → 写成 LaTeX →')
        print('        latex2omml.latex_to_omml(latex) → 替换原 w:object 那个 run')
        print('       （SKILL.md「MTEF v3 变通路线」一节有完整代码骨架）')
    print(f'清单已存: {os.path.join(outdir, "probe.json")}')


if __name__ == '__main__':
    main()
