# -*- coding: utf-8 -*-
"""
wmf_symbols — 从 MathType / 公式编辑器生成的 WMF 图元里精确提取符号

背景：老讲义里的公式大多是 MathType 的 OLE 对象，在 docx 里以 WMF 呈现图片
存着。python-docx 读不到任何文字，看起来只能 OCR。但 WMF 是**矢量**格式，
文字以 META_EXTTEXTOUT 记录原样保存，配合当时的字体（Symbol / MT Extra /
Times New Roman）就能把符号精确还原出来——比 OCR 准得多，且零误识率。

用法：
    python wmf_symbols.py <目录或单个.wmf>          # 输出符号清单
    python wmf_symbols.py <目录> --json out.json    # 导出 JSON

已知坑（都是踩过的）：
  1. WMF 头是 placeable(22B) + standard(18B) = **40 字节**，不是 22。
     算错偏移会解析出 0 条记录。
  2. LOGFONT 是 **Win16 版**（5 个 SHORT = 10B + 8 个 BYTE = 8B），
     facename 在 offset **18**，32 字节 **单字节 ANSI**，不是 UTF-16LE。
  3. MathType 输出的文字记录**按字体分组**，不是视觉顺序。
     所以只能还原「用了哪些符号」，不能还原公式结构（结构要解析 MTEF）。
"""
import argparse
import json
import os
import struct
import sys

# 只改编码，不替换 stdout 对象 —— 用 TextIOWrapper 包装会让旧 wrapper 被 GC
# 时关掉底层 buffer，两个脚本同时 import 就报 "I/O operation on closed file"
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# ---------------------------------------------------------------- 字体编码表

# Adobe Symbol 字体编码（MathType 大量用它画希腊字母和运算符）
SYMBOL = {
    0x5E: '\u22a5',   # ⊥ 垂直
    0x70: '\u03c0',   # π
    0x71: '\u03b8',   # θ 夹角
    0x72: '\u03c1',   # ρ
    0x6C: '\u03bb',   # λ
    0x6D: '\u03bc',   # μ
    0x61: '\u03b1',   # α
    0x62: '\u03b2',   # β
    0x63: '\u03c7',   # χ
    0x64: '\u03b4',   # δ
    0x65: '\u03b5',   # ε
    0x66: '\u03c6',   # φ
    0x67: '\u03b3',   # γ
    0x68: '\u03b7',   # η
    0x6E: '\u03bd',   # ν
    0x73: '\u03c3',   # σ
    0x74: '\u03c4',   # τ
    0x77: '\u03c9',   # ω
    0x78: '\u03be',   # ξ
    0x79: '\u03c8',   # ψ
    0x7A: '\u03b6',   # ζ
    0xD0: '\u2220',   # ∠ 角
    0xD6: '\u221a',   # √
    0xD7: '\u22c5',   # ⋅ 点乘
    0xB0: '\u00b0',   # °
    0xB1: '\u00b1',   # ±
    0xB3: '\u2265',   # ≥
    0xA3: '\u2264',   # ≤
    0x3D: '=', 0x2D: '\u2212', 0x2B: '+', 0x2A: '\u2217',
    0x28: '(', 0x29: ')', 0x5B: '[', 0x5D: ']', 0x7C: '|', 0x2F: '/',
    0xDB: '\u21d4',   # ⇔
    0x44: '\u0394',   # Δ 增量
    0x47: '\u03a6',   # Φ
    0x57: '\u03a9',   # Ω
    0x6A: '\u2202',   # ∂
    0xB5: '\u221e',   # ∞
    0xCE: '\u2208',   # ∈
    0xCD: '\u2209',   # ∉
    0xC8: '\u2286',   # ⊆
    0xCC: '\u2282',   # ⊂
    0x5C: '\u2234',   # ∴ 因为/所以（MathType 常用）
    0x27: '\u2032',   # ′ 分（角度）
    0x40: '\u2245',   # ≅ 全等
    0xCA: '\u2229',   # ∩ 交
    0xC9: '\u222a',   # ∪ 并
    0xCE: '\u2208',   # ∈
}

# MathType 专用字体：向量箭头靠它们拼
#   0x72 = 箭头头部；0x75 = 可重复的水平线段
#   单个 0x72        -> 短箭头（单字母向量 \vec{a}）
#   0x75×N + 0x72    -> 长箭头（多字母向量 \overrightarrow{AB}）
# 注意：MathType 6.x 之后新公式常用 **Euclid Extra** 而不是 MT Extra
# （同一批公式里两种字体可能混用），码位定义一致，必须一起识别。
MT_EXTRA_ARROW_HEAD = 0x72
MT_EXTRA_ARROW_BODY = 0x75
ARROW_FONTS = ('MT Extra', 'Euclid Extra', 'Euclid Symbol', 'Euclid Math One')


# ---------------------------------------------------------------- WMF 解析

def read_records(data):
    """切出 (record_function, body) 列表"""
    key = struct.unpack('<I', data[:4])[0]
    off = 40 if key == 0x9AC6CDD7 else 18   # placeable + standard header
    out = []
    while off + 6 <= len(data):
        recsize, recfunc = struct.unpack('<IH', data[off:off + 6])
        if recsize < 3 or off + 2 * recsize > len(data):
            break
        out.append((recfunc, data[off + 6:off + 2 * recsize]))
        off += 2 * recsize
    return out


def logfont_name(body):
    """Win16 LOGFONT -> 字体名（facename 在 offset 18，32B ANSI）"""
    return body[18:50].split(b'\x00')[0].decode('ascii', 'replace')


def decode_run(font, raw):
    """把一段字节按当时字体解码成可读符号"""
    if font == 'Symbol':
        return ''.join(SYMBOL.get(b, '[%02X]' % b) for b in raw)
    if font in ARROW_FONTS:
        n_body = sum(1 for b in raw if b == MT_EXTRA_ARROW_BODY)
        has_head = any(b == MT_EXTRA_ARROW_HEAD for b in raw)
        if has_head:
            # 箭头：3 段以上算长箭头（覆盖多字母），否则短箭头
            return ('\u2192长' if n_body >= 3 else '\u2192短') + f'(段×{n_body})'
        return ''.join('[%02X]' % b for b in raw)
    # Times New Roman / System / 其它：普通 ASCII
    return ''.join(chr(b) for b in raw)


def extract(path):
    """返回 [{'font':..., 'raw':[...], 'text':...}, ...]"""
    with open(path, 'rb') as f:
        data = f.read()
    cur = '?'
    items = []
    for func, body in read_records(data):
        if func == 0x02FB:                     # META_CREATEFONTINDIRECT
            cur = logfont_name(body)
        elif func == 0x0A32 and len(body) >= 8:  # META_EXTTEXTOUT
            slen = struct.unpack('<H', body[4:6])[0]
            fw = struct.unpack('<H', body[6:8])[0]
            rest = body[8:]
            if fw & 0x0006:                    # ETO_CLIPPED | ETO_OPAQUE
                rest = rest[8:]
            raw = rest[:slen] if slen <= len(rest) else rest
            items.append({'font': cur, 'raw': list(raw),
                          'text': decode_run(cur, raw)})
    return items


def summarize(items):
    """归纳这个图元用到了哪些关键符号"""
    flags = []
    joined = ' '.join(it['text'] for it in items)
    for it in items:
        if it['font'] in ARROW_FONTS and '\u2192' in it['text']:
            if '长' in it['text']:
                flags.append('向量(长箭头)')
            else:
                flags.append('向量(短箭头)')
    for ch, name in (('\u22a5', '垂直'), ('\u2220', '角符号'),
                     ('\u03b8', '夹角\u03b8'), ('\u03c0', '\u03c0'),
                     ('\u22c5', '点乘'), ('\u21d4', '等价')):
        if ch in joined:
            flags.append(name)
    if '//' in joined.split():
        flags.append('平行')      # MathType 用两条斜杠画 ∥
    return flags


def extract_from_docx(docx_path, outdir):
    """直接从 docx 里把 MathType 的 WMF 呈现图抽出来，返回文件路径列表

    实测结论（重要）：MathType 7.0 OLE 对象在 docx 里的**呈现图**放在
    word/media/*.wmf，而 word/embeddings/*.bin 是 OLE 复合文档本体
    （里面是 MTEF 数据，不是 WMF）。要提符号，取 media 里的 wmf。
    """
    import zipfile
    os.makedirs(outdir, exist_ok=True)
    files = []
    with zipfile.ZipFile(docx_path) as z:
        names = sorted(
            (n for n in z.namelist()
             if n.startswith('word/media/') and n.lower().endswith('.wmf')),
            key=lambda n: int(''.join(c for c in n if c.isdigit()) or 0)
        )
        for n in names:
            dst = os.path.join(outdir, n.split('/')[-1])
            with open(dst, 'wb') as f:
                f.write(z.read(n))
            files.append(dst)
    return files


def main():
    ap = argparse.ArgumentParser(description='从 MathType WMF 图元提取符号')
    ap.add_argument('target', help='.docx 文件 / 目录 / 单个 .wmf')
    ap.add_argument('--json', help='导出 JSON 到指定路径')
    ap.add_argument('--wmfdir', help='docx 模式下 WMF 的落盘目录（默认 <docx名>_wmf）')
    args = ap.parse_args()

    if os.path.isdir(args.target):
        files = sorted(
            (os.path.join(args.target, n) for n in os.listdir(args.target)
             if n.lower().endswith(('.wmf', '.emf'))),
            key=lambda p: int(''.join(c for c in os.path.basename(p) if c.isdigit()) or 0)
        )
    elif args.target.lower().endswith('.docx'):
        wdir = args.wmfdir or os.path.splitext(args.target)[0] + '_wmf'
        files = extract_from_docx(args.target, wdir)
        print(f'已从 {os.path.basename(args.target)} 抽出 {len(files)} 个 WMF -> {wdir}\n')
    else:
        files = [args.target]

    result = []
    for path in files:
        items = extract(path)
        result.append({
            'file': os.path.basename(path),
            'runs': items,
            'text': ' '.join(it['text'] for it in items),
            'symbols': summarize(items),
        })

    for r in result:
        sym = ','.join(r['symbols']) or '—'
        print(f"{r['file']:16s} {r['text']:46s} [{sym}]")

    print(f'\n共 {len(result)} 个图元')
    n_vec = sum(1 for r in result if any('向量' in s for s in r['symbols']))
    print(f'含向量符号: {n_vec}')

    import collections
    stat = collections.Counter(s for r in result for s in r['symbols'])
    if stat:
        print('\n符号分布：')
        for k, v in stat.most_common():
            print(f'  {k:14s} {v}')

    if args.json:
        with open(args.json, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=1)
        print(f'已导出 {args.json}')


if __name__ == '__main__':
    main()
