#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF -> DOCX 转换器（数学公式输出为 Word 原生可编辑公式 OMML）

思路：PDF 是矢量文本型时不走 OCR，直接由字符几何还原公式结构：
  行聚类(垂直中心) -> 上下标(字号+基线) -> 分数(分数线) -> 根号(描边路径)
  -> 定界符(可伸缩括号) -> 数学/中文分段 -> LaTeX -> OMML

用法:
  python pdf2docx_math.py <input.pdf> [-o output.docx] [--dump]
"""
import argparse
import os
import re
import sys
from collections import Counter, defaultdict

SKILL_SCRIPTS = os.path.dirname(os.path.abspath(__file__))
if SKILL_SCRIPTS not in sys.path:
    sys.path.insert(0, SKILL_SCRIPTS)

import pdfplumber                      # noqa: E402
from docx import Document              # noqa: E402
from docx.enum.table import WD_TABLE_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH   # noqa: E402
from docx.oxml import parse_xml        # noqa: E402
from docx.oxml.ns import qn           # noqa: E402
from docx.shared import Pt, RGBColor  # noqa: E402

from latex2omml import latex_to_omml  # noqa: E402


# ============================================================ Symbol 字体表
# WPS 导出的 PDF 里 SymbolMT 的 ToUnicode 把符号映射到私有区 U+F0xx，
# 需要 私有区码位 -> Symbol 原始码位 -> Unicode 的二次映射。
SYM = {
    0x20: ' ', 0x21: '!', 0x22: '∀', 0x23: '#', 0x24: '∃', 0x25: '%',
    0x26: '&', 0x27: '∋', 0x28: '(', 0x29: ')', 0x2A: '∗', 0x2B: '+',
    0x2C: ',', 0x2D: '−', 0x2E: '.', 0x2F: '/',
    0x30: '0', 0x31: '1', 0x32: '2', 0x33: '3', 0x34: '4', 0x35: '5',
    0x36: '6', 0x37: '7', 0x38: '8', 0x39: '9', 0x3A: ':', 0x3B: ';',
    0x3C: '<', 0x3D: '=', 0x3E: '>', 0x3F: '?',
    0x40: '≅', 0x41: 'Α', 0x42: 'Β', 0x43: 'Χ', 0x44: 'Δ', 0x45: 'Ε',
    0x46: 'Φ', 0x47: 'Γ', 0x48: 'Η', 0x49: 'Ι', 0x4A: 'ϑ', 0x4B: 'Κ',
    0x4C: 'Λ', 0x4D: 'Μ', 0x4E: 'Ν', 0x4F: 'Ο',
    0x50: 'Π', 0x51: 'Θ', 0x52: 'Ρ', 0x53: 'Σ', 0x54: 'Τ', 0x55: 'Υ',
    0x56: 'ς', 0x57: 'Ω', 0x58: 'Ξ', 0x59: 'Ψ', 0x5A: 'Ζ',
    0x5B: '[', 0x5C: '∴', 0x5D: ']', 0x5E: '⊥', 0x5F: '_',
    0x60: '‾', 0x61: 'α', 0x62: 'β', 0x63: 'χ', 0x64: 'δ', 0x65: 'ε',
    0x66: 'φ', 0x67: 'γ', 0x68: 'η', 0x69: 'ι', 0x6A: 'ϕ', 0x6B: 'κ',
    0x6C: 'λ', 0x6D: 'μ', 0x6E: 'ν', 0x6F: 'ο',
    0x70: 'π', 0x71: 'θ', 0x72: 'ρ', 0x73: 'σ', 0x74: 'τ', 0x75: 'υ',
    0x76: 'ϖ', 0x77: 'ω', 0x78: 'ξ', 0x79: 'ψ', 0x7A: 'ζ',
    0x7B: '{', 0x7C: '|', 0x7D: '}', 0x7E: '~',
    0xA0: '€', 0xA1: 'ϒ', 0xA2: '′', 0xA3: '≤', 0xA4: '⁄', 0xA5: '∞',
    0xA6: 'ƒ', 0xA7: '♣', 0xA8: '♦', 0xA9: '♥', 0xAA: '♠',
    0xAB: '↔', 0xAC: '←', 0xAD: '↑', 0xAE: '→', 0xAF: '↓',
    0xB0: '°', 0xB1: '±', 0xB2: '″', 0xB3: '≥', 0xB4: '×', 0xB5: '∝',
    0xB6: '∂', 0xB7: '•', 0xB8: '÷', 0xB9: '≠', 0xBA: '≡', 0xBB: '≈',
    0xBC: '…', 0xBD: '⏐', 0xBE: '⌈', 0xBF: '⌉', 0xC0: '⌊', 0xC1: '⌋',
    0xC2: '⍀', 0xC3: '⟨', 0xC4: '⌠', 0xC5: '⎧', 0xC6: '⌡', 0xC7: '⎨',
    0xC8: '⎩', 0xC9: '⎪', 0xCA: '⟩', 0xCB: '⊃', 0xCC: '∩', 0xCD: '⊂',
    0xCE: '∈', 0xCF: '∉', 0xD0: '∠', 0xD1: '∇', 0xD2: 'Ⓡ', 0xD3: '©',
    0xD4: '™', 0xD5: '∏', 0xD6: '√', 0xD7: '⋅', 0xD8: '¬', 0xD9: '∧',
    0xDA: '∨', 0xDB: '⇔', 0xDC: '⇐', 0xDD: '⇑', 0xDE: '⇒', 0xDF: '⇓',
    0xE0: '◊', 0xE1: '〈', 0xE2: '®', 0xE3: '©', 0xE4: '™', 0xE5: '∑',
    0xE6: '⎛', 0xE7: '⎜', 0xE8: '⎝', 0xE9: '⎡', 0xEA: '⎢', 0xEB: '⎣',
    0xEC: '⎧', 0xED: '⎨', 0xEE: '⎩', 0xEF: '⎪',
    0xF0: '⎫', 0xF1: '⎬', 0xF2: '⎭', 0xF3: '⎮', 0xF4: '⎞', 0xF5: '⎟',
    0xF6: '⎠', 0xF7: '⎤', 0xF8: '⎥', 0xF9: '⎦', 0xFA: '⎱', 0xFB: '⎰',
    0xFC: '⎲', 0xFD: '⎳', 0xFE: '⎴', 0xFF: '',
}

# 可伸缩定界符的拼接件（不是内容，是括号的上下半/中段）
DELIM_PIECES = set('⎛⎜⎝⎞⎟⎠⎡⎢⎣⎤⎥⎦⎧⎨⎩⎪⎫⎬⎭⎮⎱⎰⌠⌡⏐'
                   '∣⌈⌉⌊⌋〈〉')

# WPS 伸缩小括号的拼接件 ToUnicode 常为空文本，只能靠原始 Symbol 码位分类
PIECE_CLASS = {}
for _c in list(range(0xE6, 0xE9)) + list(range(0xF4, 0xF7)):
    PIECE_CLASS[_c] = 'paren'
for _c in list(range(0xE9, 0xEC)) + list(range(0xF7, 0xFA)):
    PIECE_CLASS[_c] = 'bracket'
for _c in list(range(0xEC, 0xEF)) + list(range(0xF0, 0xF3)):
    PIECE_CLASS[_c] = 'brace'
for _c in (0xFA, 0xFB):
    PIECE_CLASS[_c] = 'bracket'

FENCE_TEX = {
    'paren': (r'\left( ', r'\,\right)'),
    'bracket': (r'\left[ ', r'\,\right]'),
    'brace': (r'\left\{ ', r'\,\right\}'),
}


def remap_symbol(text):
    """SymbolMT 私有区码位 -> Unicode。非私有区原样返回。"""
    if len(text) != 1:
        return text
    o = ord(text)
    if 0xF000 <= o <= 0xF0FF:
        return SYM.get(o - 0xF000, '')
    return text


# ============================================================ Unicode -> LaTeX
_GREEK_L = ([chr(c) for c in range(0x3B1, 0x3C2)] +
            [chr(c) for c in range(0x3C3, 0x3CA)])
_GREEK_U = ([chr(c) for c in range(0x391, 0x3A2)] +
            [chr(c) for c in range(0x3A3, 0x3AA)])
_GREEK_N = ['alpha', 'beta', 'gamma', 'delta', 'epsilon', 'zeta', 'eta',
            'theta', 'iota', 'kappa', 'lambda', 'mu', 'nu', 'xi', 'omicron',
            'pi', 'rho', 'sigma', 'tau', 'upsilon', 'phi', 'chi', 'psi',
            'omega']

UNI2TEX = {
    '∀': r'\forall', '∃': r'\exists', '∈': r'\in', '∉': r'\notin',
    '≤': r'\le', '≥': r'\ge', '≠': r'\ne', '¬': r'\neg', '∞': r'\infty',
    '×': r'\times', '÷': r'\div', '±': r'\pm', '⋅': r'\cdot', '•': r'\cdot',
    '√': r'\sqrt', '∝': r'\propto', '≡': r'\equiv', '≈': r'\approx',
    '⊂': r'\subset', '⊃': r'\supset', '⊆': r'\subseteq', '⊇': r'\supseteq',
    '∪': r'\cup', '∩': r'\cap', '∅': r'\varnothing', '∅': r'\varnothing',
    '→': r'\to', '←': r'\leftarrow', '⇒': r'\Rightarrow',
    '⇐': r'\Leftarrow', '⇔': r'\Leftrightarrow', '↔': r'\leftrightarrow',
    '⊥': r'\perp', '∥': r'\parallel', '∠': r'\angle', '∇': r'\nabla',
    '∂': r'\partial', '∑': r'\sum', '∏': r'\prod', '∫': r'\int',
    '…': r'\ldots', '⋯': r'\cdots', '′': "'", '″': "''", '°': r'^{\circ}',
    '∗': '*', '⁄': '/', '−': '-', '∠': r'\angle',
    '⌈': r'\lceil', '⌉': r'\rceil', '⌊': r'\lfloor', '⌋': r'\rfloor',
    '⟨': r'\langle', '⟩': r'\rangle',
    'ς': r'\varsigma', 'ϑ': r'\vartheta', 'ϕ': r'\varphi', 'ϖ': r'\varpi',
    '≅': r'\cong', '∋': r'\ni', '⏐': '|',
    # --- 线性代数
    '⊗': r'\otimes', '⊕': r'\oplus', '⊖': r'\ominus', '⊘': r'\oslash',
    '‖': r'\Vert',
    # --- 数论
    '∣': r'\mid', '∤': r'\nmid', '∦': r'\nparallel',
    'ℂ': r'\mathbb{C}', 'ℍ': r'\mathbb{H}', 'ℕ': r'\mathbb{N}',
    'ℙ': r'\mathbb{P}', 'ℚ': r'\mathbb{Q}', 'ℝ': r'\mathbb{R}',
    'ℤ': r'\mathbb{Z}',
    # --- 平面几何
    '△': r'\triangle', '∆': r'\Delta', '⊙': r'\odot', '⌒': r'\frown',
    '∽': r'\backsim', '∡': r'\measuredangle',
    # --- 统计
    '∼': r'\sim',
    # --- 单位/物理
    '℃': r'^{\circ}\mathrm{C}', '℉': r'^{\circ}\mathrm{F}',
    '‰': r'\text{‰}', 'Å': r'\text{Å}', 'Ω': r'\Omega',
    'µ': r'\mu', '″': "''",
    # --- 微积分/场论
    '∂': r'\partial', '∇': r'\nabla',
    '∬': r'\iint', '∭': r'\iiint', '∮': r'\oint',
    '∯': r'\oiint', '∰': r'\oiiint', '⨌': r'\int\!\!\!\!\int',
    '∇': r'\nabla',
    '⇌': r'\rightleftharpoons', '⇋': r'\leftrightharpoons',
    '⋁': r'\bigvee', '⋀': r'\bigwedge', '⋂': r'\bigcap', '⋃': r'\bigcup',
    # --- 化学
    '↑': r'\uparrow', '↓': r'\downarrow',
    '·': r'\cdot', '⟶': r'\longrightarrow',
    '⇀': r'\rightharpoonup', '↽': r'\leftharpoondown',
    '←': r'\leftarrow',
    # --- 生物/地理
    '♀': r'\female', '♂': r'\male',
    '‱': r'\text{‱}', '≈': r'\approx', '∝': r'\propto',
}
for _c, _n in zip(_GREEK_L, _GREEK_N):
    UNI2TEX[_c] = '\\' + _n
for _c, _n in zip(_GREEK_U, _GREEK_N):
    UNI2TEX[_c] = '\\' + _n.capitalize()
# Ohm 符号/微符号的兼容码位（Ω 有 U+03A9 与 U+2126 两个码位，µ 同理）
UNI2TEX['\u2126'] = r'\Omega'
UNI2TEX['\u00b5'] = r'\mu'

# 数学字母数字变体（U+1D400–1D7FF）：Word/WPS 导出的 PDF 常用斜体/粗体/双线体
# 数学字母（𝑥 𝐴 ℝ 𝒩…），PDF 提取后是这些变体码位而非普通字母。
# 全部还原成普通字母，格式差异由字体呈现层处理，不进公式语义。
def _register_math_alphanumerics():
    # (起始码位, 起始大写字母, 大写 26 个 + 小写 26 个)
    blocks = [
        (0x1D400, 'A', 'bold'),
        (0x1D434, 'A', 'italic'),
        (0x1D468, 'A', 'bold-italic'),
        (0x1D49C, 'A', 'script'),
        (0x1D4D0, 'A', 'bold-script'),
        (0x1D504, 'A', 'fraktur'),
        (0x1D538, 'A', 'double-struck'),
        (0x1D56C, 'A', 'bold-fraktur'),
        (0x1D5A0, 'A', 'sans'),
        (0x1D5D4, 'A', 'sans-bold'),
        (0x1D608, 'A', 'sans-italic'),
        (0x1D63C, 'A', 'sans-bold-italic'),
        (0x1D670, 'A', 'monospace'),
    ]
    holes = {  # 该块内被保留字符占用的缺口（Unicode 标准）
        0x1D49C: [0x1D49D, 0x1D4A0, 0x1D4A1, 0x1D4A3, 0x1D4A4, 0x1D4A7,
                  0x1D4A8, 0x1D4AD],
        0x1D504: [0x1D506, 0x1D50B, 0x1D50C, 0x1D515, 0x1D51D],
        0x1D538: [0x1D53A, 0x1D53F, 0x1D545, 0x1D547, 0x1D548, 0x1D549],
    }
    for start, base, _style in blocks:
        for i in range(52):                      # 26 大写 + 26 小写
            cp = start + i
            if cp in holes.get(start, ()):
                continue
            if i < 26:
                ch = chr(ord(base) + i)
            else:
                ch = chr(ord('a') + i - 26)
            UNI2TEX.setdefault(chr(cp), ch)
    # letterlike 双线体（不在 1D538 块内的 6 个）
    for cp, name in ((0x2102, 'C'), (0x210D, 'H'), (0x2115, 'N'),
                     (0x2119, 'P'), (0x211A, 'Q'), (0x211D, 'R'),
                     (0x2124, 'Z')):
        UNI2TEX[chr(cp)] = r'\mathbb{%s}' % name


_register_math_alphanumerics()

# 需要转义的 ASCII
_ESCAPE = {'{': r'\{', '}': r'\}', '|': r'\mid', '&': r'\&', '%': r'\%',
           '$': r'\$', '#': r'\#', '_': r'\_', '~': r'\sim'}

# 出现这些字符的数学串必须包进 $...$
_FORCE_MATH = set('∀∃∈∉≤≥≠¬∞√×÷±⊥∥∠⊆⊇⊂⊃∪∩→⇒⇔∑∏∫′″∗⁄⋅'
                  '⊗⊕⊖⊘‖∣∤∦△∆⊙⌒∼∽∡ℂℍℕℙℚℝℤ‰'
                  '∂∇∬∭∮∯∰⨌⇌⇋⋁⋀⋂⋃↑↓⇑⇓↕'
                  '←⟶⇀↽°♀♂‱≈∝')


def tex_atom(ch):
    if ch in UNI2TEX:
        return UNI2TEX[ch]
    if ch in _ESCAPE:
        return _ESCAPE[ch]
    return ch


# ============================================================ 提取
def extract(path):
    """返回每页 dict: {page, width, height, chars, lines, rects, curves}"""
    pages = []
    with pdfplumber.open(path) as pdf:
        for i, pg in enumerate(pdf.pages, 1):
            chars = []
            for c in pg.chars:
                font = c.get('fontname', '')
                text = c.get('text') or ''
                raw = text
                if 'Symbol' in font:
                    text = remap_symbol(text)
                chars.append({
                    'x0': c['x0'], 'x1': c['x1'],
                    'top': c['top'], 'bottom': c['bottom'],
                    'size': c['size'], 'font': font, 'text': text,
                    'raw': raw, 'page': i,
                })
            segs = []
            for src in ('lines', 'curves'):
                for o in getattr(pg, src, []):
                    segs.append({
                        'x0': o['x0'], 'x1': o['x1'],
                        'top': o['top'], 'bottom': o['bottom'],
                        'lw': o.get('linewidth', 0),
                    })
            pages.append({
                'page': i, 'width': float(pg.width), 'height': float(pg.height),
                'chars': chars, 'segs': segs,
                'rects': [{'x0': r['x0'], 'x1': r['x1'], 'top': r['top'],
                           'bottom': r['bottom']} for r in pg.rects],
            })
    return pages


def ctr(c):
    return (c['top'] + c['bottom']) / 2.0


# ============================================================ 结构识别
def detect_radicals(page):
    """根号：一段水平描边（顶上的横线）+ 左侧向下勾。返回 (x0, x1, y)"""
    out = []
    horiz = []
    for s in page['segs']:
        if abs(s['top'] - s['bottom']) < 0.4 and (s['x1'] - s['x0']) >= 6:
            horiz.append(s)
    for h in horiz:
        hook = None
        for s in page['segs']:
            if s is h:
                continue
            if abs(s['x1'] - h['x0']) < 0.6 and s['bottom'] > h['top'] + 3 \
                    and s['top'] < h['top'] + 2.0:
                hook = s
                break
        if hook is not None:
            out.append((h['x0'], h['x1'], h['top']))
    return out


def detect_tables(page):
    """由长网格线还原表格。返回 [{x0,x1,ybounds,colbounds}]"""
    hs, vs = [], []
    for s in page['segs']:
        if abs(s['top'] - s['bottom']) < 0.4 and (s['x1'] - s['x0']) > 60:
            hs.append(s)
        elif abs(s['x1'] - s['x0']) < 0.4 and (s['bottom'] - s['top']) > 60:
            vs.append(s)
    groups = defaultdict(list)
    for h in hs:
        groups[(round(h['x0'], 0), round(h['x1'], 0))].append(h)
    tables = []
    for (gx0, gx1), hs_ in groups.items():
        if len(hs_) < 3:
            continue
        ys = sorted(set(round(h['top'], 1) for h in hs_))
        y0, y1 = ys[0], ys[-1]
        cols = sorted(set(round(v['x0'], 1) for v in vs
                          if gx0 - 2 <= v['x0'] <= gx1 + 2
                          and v['top'] <= y0 + 2 and v['bottom'] >= y1 - 2))
        if len(cols) < 2:
            continue
        tables.append({'x0': min(cols), 'x1': max(cols),
                       'ys': ys, 'cols': cols})
    tables.sort(key=lambda t: (t['ys'][0], t['x0']))
    return tables


def detect_fractions(page, chars):
    """分数线：短水平线，上方有字符、下方有字符。返回 [(num, den, x, y)]"""
    out = []
    for s in page['segs']:
        if abs(s['top'] - s['bottom']) > 0.4:
            continue
        w = s['x1'] - s['x0']
        if not (3 <= w <= 60):
            continue
        if s.get('lw', 0) <= 0:
            continue
        y = s['top']
        num = [c for c in chars
               if s['x0'] - 2 <= c['x0'] <= s['x1'] + 2
               and y - 15 <= c['bottom'] <= y - 0.2
               and c['text'].strip()]
        den = [c for c in chars
               if s['x0'] - 2 <= c['x0'] <= s['x1'] + 2
               and y + 0.2 <= c['top'] <= y + 15
               and c['text'].strip()]
        if num and den:
            out.append({'num': num, 'den': den,
                        'x': (s['x0'] + s['x1']) / 2, 'y': y,
                        'x0': s['x0'], 'x1': s['x1']})
    return out


def detect_fillins(page, fractions):
    """填空横线：与内容同行的短横线（非分数线、非表格线）"""
    out = []
    for s in page['segs']:
        if abs(s['top'] - s['bottom']) > 0.4:
            continue
        w = s['x1'] - s['x0']
        if not (20 <= w <= 90) or s.get('lw', 0) <= 0:
            continue
        if fractions and any(abs(f['x0'] - s['x0']) < 1 and
                             abs(f['y'] - s['top']) < 1 for f in fractions):
            continue
        out.append({'x': (s['x0'] + s['x1']) / 2, 'y': s['top'],
                    'x0': s['x0'], 'x1': s['x1'], 'w': w})
    return out


def detect_bars(page):
    """矢量竖线：集合中 {x | ...} 的分隔竖杠常被画成描边而非字符"""
    out = []
    for s in page['segs']:
        if abs(s['x1'] - s['x0']) > 0.4 or s.get('lw', 0) <= 0:
            continue
        h = s['bottom'] - s['top']
        if 6 <= h <= 40:
            out.append({'x': s['x0'], 'y0': s['top'], 'y1': s['bottom']})
    return out


def _piece_class(c):
    """单个字符若是定界符拼接件，返回类型；否则 None。
    raw 码位是 U+F0xx 私有区，需先减 0xF000。"""
    t = c['text']
    raw = c.get('raw') or t
    code = None
    if len(raw) == 1:
        o = ord(raw)
        code = o - 0xF000 if 0xF000 <= o <= 0xF0FF else o
    if code is not None and code in PIECE_CLASS:
        if (len(t) == 1 and t in DELIM_PIECES) or not t.strip():
            return PIECE_CLASS[code]
    return None


def detect_fences(chars):
    """可伸缩定界符：同一 x 上纵向连续堆叠的拼接件。
    返回 (成对括号, 未配对大括号)。空 ToUnicode 的拼接件靠原始码位识别。"""
    cols = defaultdict(list)
    for c in chars:
        cls = _piece_class(c)
        if cls is not None:
            cols[round(c['x0'], 1)].append((c, cls))
    stacks = []
    for x, items in cols.items():
        if not items:
            continue
        items.sort(key=lambda p: p[0]['top'])
        # 同一列可能是多个独立括号（如两个大括号组），按纵向连续性切段
        runs = [[items[0]]]
        for it in items[1:]:
            if it[0]['top'] - runs[-1][-1][0]['bottom'] > 4:
                runs.append([it])
            else:
                runs[-1].append(it)
        for run in runs:
            if len(run) < 2:
                continue
            classes = [cls for _, cls in run if cls]
            typ = Counter(classes).most_common(1)[0][0] if classes else None
            stacks.append({'x': x,
                           'y0': min(c['top'] for c, _ in run),
                           'y1': max(c['bottom'] for c, _ in run),
                           'n': len(run), 'type': typ,
                           'chars': [c for c, _ in run]})
    fences = []
    used = set()
    for i, a in enumerate(stacks):
        if i in used or a['type'] is None:
            continue
        for j, b in enumerate(stacks):
            if j <= i or j in used or b['type'] != a['type']:
                continue
            if abs(a['y0'] - b['y0']) < 2.5 and abs(a['y1'] - b['y1']) < 2.5:
                lo, hi = (a, b) if a['x'] < b['x'] else (b, a)
                if 12 <= (hi['x'] - lo['x']) <= 90:
                    fences.append({'x0': lo['x'], 'x1': hi['x'],
                                   'y0': a['y0'], 'y1': a['y1'],
                                   'type': a['type'],
                                   'chars': lo['chars'] + hi['chars']})
                    used.add(i)
                    used.add(j)
                break
    braces = [s for k, s in enumerate(stacks)
              if k not in used and s['type'] == 'brace']
    return fences, braces


# ============================================================ 行聚类
def cluster_rows(chars, ref, tol=5.0):
    """按垂直中心聚类成行；上下标不参与锚定，按基线就近归行。"""
    def is_script(c):
        return c['size'] < ref * 0.75

    anchors = [c for c in chars if not is_script(c) and c['text'].strip()]
    scripts = [c for c in chars if is_script(c) and c['text'].strip()]
    if not anchors:
        return []
    anchors.sort(key=lambda c: (ctr(c), c['x0']))
    rows = [[anchors[0]]]
    for c in anchors[1:]:
        if ctr(c) - ctr(rows[-1][0]) > tol:
            rows.append([c])
        else:
            rows[-1].append(c)

    # 行基线 = 该行字符 bottom 的众数
    for r in rows:
        bots = [round(c['bottom'], 1) for c in r]
        r.sort(key=lambda c: c['x0'])
        rows[rows.index(r)] = r
    baselines = [Counter(round(c['bottom'], 1) for c in r).most_common(1)[0][0]
                 for r in rows]

    for c in scripts:
        best, bd = None, 1e9
        for k, r in enumerate(rows):
            d = abs(c['bottom'] - baselines[k]) if c['bottom'] < baselines[k] \
                else abs(c['top'] - baselines[k])
            if d < bd:
                bd, best = d, k
        if best is not None and bd < 16:
            rows[best].append(c)
    for r in rows:
        r.sort(key=lambda c: c['x0'])
    return rows


# ============================================================ 行 -> 片段
def chars_tex(chars, ref):
    """一小段字符 -> LaTeX（用于分数分子/分母）"""
    return ''.join(atom_tex(a) for a in row_atoms(chars, ref))


def row_atoms(row, ref):
    """行内字符 -> atom 列表 [{x, tex, ismath, base}]，上下标已并入主体。"""
    bots = [round(c['bottom'], 1) for c in row if c['size'] >= ref * 0.75]
    baseline = Counter(bots).most_common(1)[0][0] if bots else \
        Counter(round(c['bottom'], 1) for c in row).most_common(1)[0][0]

    atoms = []
    for c in row:
        if c.get('_frac'):
            f = c['_frac']
            atoms.append({
                'x': c['x0'], 'ch': '', 'base': c, 'math': True,
                'tex': r'\frac{%s}{%s}' % (chars_tex(f['num'], ref),
                                           chars_tex(f['den'], ref)),
            })
            continue
        if not c['text'].strip():
            continue
        is_sup = is_sub = False
        if c['size'] < ref * 0.75:
            # 判据用「底边 vs 基线」：底边明显低于基线=下标，明显高于=上标
            if c['bottom'] <= baseline - 1.0:
                is_sup = True
            elif c['bottom'] >= baseline + 0.5:
                is_sub = True
            else:
                is_sup = ctr(c) < baseline - 1.0
        if is_sup or is_sub:
            if atoms:
                atoms[-1].setdefault('sup' if is_sup else 'sub', []).append(c)
            continue
        atoms.append({'x': c['x0'], 'ch': c['text'], 'base': c})
    return atoms


def atom_tex(a):
    body = tex_atom(a['ch'])
    # 命令名后补空格，避免 \alpha 与后续字母粘连成 \alphax
    if body.startswith('\\') and body[-1:].isalpha():
        body += ' '
    for key, sym in (('sub', '_'), ('sup', '^')):
        if a.get(key):
            inner = ''.join(tex_atom(c['text']) for c in a[key])
            body += sym + '{' + inner + '}'
    return body


def is_math_char(c):
    f = c['font']
    t = c['text']
    if t in _FORCE_MATH:
        return True
    if 'Symbol' in f:
        return True
    if 'Cambria' in f or 'Euclid' in f or 'Math' in f:
        return True
    if 'Times' in f:
        return bool(t) and (t.isascii() and (t.isalnum() or t in '.,:;()[]+-*/=<>|'))
    if 'Malgun' in f and t.isascii() and t.isalnum():
        return True
    return False


def wrap_range(atoms, x0, x1, open_tex, close_tex):
    """把 x 落在 [x0,x1] 的连续 atom 用 open...close 包裹。
    若末位 atom 的下标其实落在定界符外侧（如 [f(x)]_min），把下标移到括号外。"""
    idx = [i for i, a in enumerate(atoms)
           if x0 - 1 <= a['x'] <= x1 + 1 and a.get('math', False)]
    if not idx:
        return
    lo, hi = min(idx), max(idx)
    inner = ''.join(a['tex'] for a in atoms[lo:hi + 1])
    tail = ''
    subs = atoms[hi].get('sub') or []
    if subs and all(c['x0'] > x1 + 1 for c in subs):
        cut = len('_{' + ''.join(tex_atom(c['text']) for c in subs) + '}')
        if inner.endswith('_{' + ''.join(tex_atom(c['text'])
                                         for c in subs) + '}'):
            inner, tail = inner[:-cut], inner[-cut:]
    atoms[lo:hi + 1] = [{'x': atoms[lo]['x'], 'math': True, 'ch': '',
                         'tex': open_tex + inner + close_tex + tail}]


def row_segments(row, ref, ctx):
    """行 -> [(kind, value)]，kind ∈ text/math"""
    atoms = row_atoms(row, ref)
    for a in atoms:
        if 'tex' not in a:
            a['tex'] = atom_tex(a)
        a['math'] = a.get('math', False) or is_math_char(a['base'])
    ytop = min(c['top'] for c in row)
    ybot = max(c['bottom'] for c in row)
    for rad in ctx['radicals']:
        if ytop - 4 <= rad[2] <= ybot + 4:
            wrap_range(atoms, rad[0], rad[1], r'\sqrt{', '}')
    for f in ctx['fences']:
        if abs(f['y0'] - ytop) < 8 and abs(f['y1'] - ybot) < 10:
            o, cl = FENCE_TEX.get(f.get('type', 'bracket'),
                                  (r'\left[ ', r'\,\right]'))
            wrap_range(atoms, f['x0'], f['x1'], o, cl)

    segs = []
    buf, bufmath = [], None
    for a in atoms:
        m = a.get('math', False) or any(ch in _FORCE_MATH for ch in a['tex'])
        if bufmath is None:
            bufmath = m
        if m != bufmath and buf:
            segs.append((('math' if bufmath else 'text'), ''.join(buf)))
            buf, bufmath = [], m
        buf.append(a['tex'] if m else a['ch'])
    if buf:
        segs.append((('math' if bufmath else 'text'), ''.join(buf)))

    # 数学段里若既没有字母也没有必需 LaTeX 的符号（如 "1.5"、"2"），当正文处理
    out = []
    for kind, val in segs:
        if kind == 'math' and not any(ch.isalpha() for ch in val) \
                and not any(ch in _FORCE_MATH for ch in val):
            out.append(('text', val))
        else:
            out.append((kind, val))
    return out


# ============================================================ DOCX 写出
M_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def set_math_size(om_el, msize):
    """给 OMML 公式的每个文本 run 设字号（半点单位），对齐 PDF 原始字号。"""
    sz = str(int(round(msize * 2)))
    for r in om_el.iter(M_NS + 'r'):
        rpr = r.find(qn('w:rPr'))
        if rpr is None:
            rpr = r.makeelement(qn('w:rPr'), {})
            # w:rPr 必须在 m:t 之前；若已有 m:rPr 则排其后
            idx = 0
            if len(r) and r[0].tag == M_NS + 'rPr':
                idx = 1
            r.insert(idx, rpr)
        for tagname in ('w:sz', 'w:szCs'):
            e = rpr.find(qn(tagname))
            if e is None:
                e = rpr.makeelement(qn(tagname), {})
                rpr.append(e)
            e.set(qn('w:val'), sz)


def style_run(r, size, bold=False):
    """正文 run：黑色 + Times New Roman / 微软雅黑（与 PDF 一致）"""
    r.font.name = 'Times New Roman'
    r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    r.font.size = Pt(size)
    r.font.color.rgb = RGBColor(0, 0, 0)
    if bold:
        r.bold = True


def write_para(p, segs, bold=False, tsize=10.5, msize=12.0):
    """把 [(text|math, value)] 写进已有段落；math 段转成原生 OMML 公式"""
    for kind, val in segs:
        if not val:
            continue
        if kind == 'text':
            # 填空横线用「空格 + 下划线格式」呈现，而不是下划线字符
            for part in re.split('(＿+)', val):
                if not part:
                    continue
                if part[0] == '＿':
                    r = p.add_run('\u3000' * len(part))
                    r.underline = True
                    style_run(r, tsize, bold)
                else:
                    r = p.add_run(part)
                    style_run(r, tsize, bold)
        else:
            try:
                el = parse_xml(latex_to_omml(val, 'inline'))
                set_math_size(el, msize)
                p._p.append(el)
            except Exception as e:                       # noqa: BLE001
                r = p.add_run('[%s]' % val)
                style_run(r, tsize, bold)
                sys.stderr.write('[warn] 公式转换失败 %r -> %s\n' % (val, e))
    return p


def add_rich(doc, segs, style=None, align=None, indent_pt=0.0, bold=False,
             tsize=10.5, msize=12.0):
    p = doc.add_paragraph(style=style)
    if align is not None:
        p.alignment = align
    if indent_pt:
        p.paragraph_format.left_indent = Pt(indent_pt)
    return write_para(p, segs, bold=bold, tsize=tsize, msize=msize)


def brace_group_xml(segs_list, msize):
    """多行条件组 -> OMML：m:d(begChr={, endChr=) + m:eqArr，Word 原生方程数组"""
    from xml.sax.saxutils import escape
    sz = str(int(round(msize * 2)))

    def run_xml(kind, val):
        if kind == 'math':
            om = latex_to_omml(val, 'inline')
            if isinstance(om, bytes):
                om = om.decode('utf-8')
            om = om.strip()
            return om[om.index('>') + 1: om.rindex('</m:oMath>')]
        return ('<m:r><w:rPr><w:rFonts w:ascii="Times New Roman"'
                ' w:hAnsi="Times New Roman" w:eastAsia="微软雅黑"/>'
                '<w:color w:val="000000"/><w:sz w:val="%s"/></w:rPr>'
                '<m:t xml:space="preserve">%s</m:t></m:r>'
                ) % (sz, escape(val))

    rows_xml = ''.join(
        '<m:e>%s</m:e>' % ''.join(run_xml(k, v) for k, v in segs if v)
        for segs in segs_list)
    return (
        '<m:oMath xmlns:m="%s" xmlns:w="%s"><m:d><m:dPr>'
        '<m:begChr m:val="{"/><m:endChr m:val=""/>'
        '<m:ctrlPr><w:rPr><w:sz w:val="%s"/></w:rPr></m:ctrlPr>'
        '</m:dPr><m:e><m:eqArr>%s</m:eqArr></m:e></m:d></m:oMath>'
    ) % ('http://schemas.openxmlformats.org/officeDocument/2006/math',
         W_NS, sz, rows_xml)


def add_brace_group(doc, segs_list, indent_pt, msize, label_segs=None):
    p = doc.add_paragraph()
    if label_segs:
        write_para(p, [s for segs in label_segs for s in segs
                      if s[0] == 'text' and s[1].strip()],
                   tsize=10.5, msize=msize)
    if indent_pt and not label_segs:
        p.paragraph_format.left_indent = Pt(indent_pt)
    p._p.append(parse_xml(brace_group_xml(segs_list, msize)))
    for om in p._p.findall('.//' + M_NS + 'oMath'):
        set_math_size(om, msize)
    return p


# ============================================================ 主流程
HEAD_RE = re.compile(r'^[一二三四五六七八九十]+、')
KEY_RE = re.compile(r'^(例\s*\d|变式\s*\d|复习思考|课堂引入|实例|规律方法|注|总结)[：:]')


def row_tsize(r, ref):
    """行的正文字号（非数学字符的众数）"""
    ts = [round(c['size'], 1) for c in r
          if c['text'].strip() and not is_math_char(c)]
    return Counter(ts).most_common(1)[0][0] if ts else ref


def row_msize(r, ref):
    """行的公式字号（数学字符的众数），PDF 里公式比正文大一号"""
    ms = [round(c['size'], 1) for c in r
          if c['text'].strip() and c['size'] >= ref * 0.75
          and is_math_char(c)]
    return Counter(ms).most_common(1)[0][0] if ms else 12.0


def convert(pdf_path, out_path, dump=False):
    pages = extract(pdf_path)
    doc = Document()
    st = doc.styles['Normal']
    st.font.name = 'Times New Roman'
    st.font.size = Pt(10.5)
    st.font.color.rgb = RGBColor(0, 0, 0)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    # 版心对齐 PDF：PDF 版心宽 ~420pt，Word 默认 3.17cm 边距只有 415.6pt，
    # 长公式会跨行；收窄到 80pt(2.82cm) 留出余量
    LEFT_PT = 80.0
    sec = doc.sections[0]
    sec.left_margin = Pt(LEFT_PT)
    sec.right_margin = Pt(LEFT_PT)

    total_math, failed = 0, 0

    for pi, page in enumerate(pages):
        chars = page['chars']
        ref = Counter(round(c['size'], 1) for c in chars).most_common(1)[0][0]

        radicals = detect_radicals(page)
        tables = detect_tables(page)
        bars = detect_bars(page)

        # 表格占用的字符
        in_tbl = set()
        tbl_cells = []
        for t in tables:
            nrow = len(t['ys']) - 1
            ncol = len(t['cols']) - 1
            grid = [[[] for _ in range(ncol)] for _ in range(nrow)]
            for c in chars:
                if not (t['x0'] - 1 <= c['x0'] <= t['x1'] + 1):
                    continue
                cx = ctr(c)
                ri = None
                for k in range(nrow):
                    if t['ys'][k] - 1 <= cx <= t['ys'][k + 1] + 1:
                        ri = k
                        break
                ci = None
                for k in range(ncol):
                    if t['cols'][k] - 1 <= c['x0'] <= t['cols'][k + 1] + 1:
                        ci = k
                        break
                if ri is not None and ci is not None:
                    grid[ri][ci].append(c)
                    in_tbl.add(id(c))
            tbl_cells.append({'y': t['ys'][0], 'nrow': nrow, 'ncol': ncol,
                              'grid': grid, 'x0': t['x0']})

        free = [c for c in chars if id(c) not in in_tbl]
        frac = detect_fractions(page, free)
        fences, braces = detect_fences(free)

        # 分数：把分子分母从原行摘掉，作为一个伪字符插回分数线所在行
        frac_ids = set()
        for f in frac:
            for c in f['num'] + f['den']:
                frac_ids.add(id(c))
        # 可伸缩定界符的拼接件不是内容，全部剔除
        delim_ids = set()
        for grp in fences:
            for c in grp['chars']:
                delim_ids.add(id(c))
        for b in braces:
            for c in b['chars']:
                delim_ids.add(id(c))
        for c in free:
            if len(c['text']) == 1 and c['text'] in DELIM_PIECES:
                delim_ids.add(id(c))
        free = [c for c in free
                if id(c) not in frac_ids and id(c) not in delim_ids]
        fillins = detect_fillins(page, frac)

        ctx = {'radicals': radicals, 'fences': fences}
        left = min((c['x0'] for c in free), default=90.0)

        rows = cluster_rows(free, ref)
        items = []
        row_segs = {}

        # 每个分数只归到离分数线最近的那一行
        row_mid = []
        for r in rows:
            row_mid.append(sum(ctr(c) for c in r) / len(r))
        frac_of_row = defaultdict(list)
        for f in frac:
            if not rows:
                continue
            k = min(range(len(rows)), key=lambda i: abs(row_mid[i] - f['y']))
            frac_of_row[k].append(f)

        for k, r in enumerate(rows):
            rowc = list(r)
            for f in frac_of_row[k]:
                rowc.append({'x0': f['x'], 'x1': f['x'],
                             'top': f['y'] - 1, 'bottom': f['y'] + 1,
                             'size': ref, 'font': 'TimesNewRomanPSMT',
                             'text': '', 'page': page['page'], '_frac': f})
            # 矢量竖杠 -> \mid
            for b in bars:
                if b['y0'] - 10 <= row_mid[k] <= b['y1'] + 10:
                    rowc.append({'x0': b['x'], 'x1': b['x'],
                                 'top': b['y0'], 'bottom': b['y1'],
                                 'size': ref, 'font': 'TimesNewRomanPSMT',
                                 'text': '|', 'page': page['page']})
            # 填空横线 -> 下划线文本，按 x 插回行内
            for fl in fillins:
                if abs(fl['y'] - row_mid[k]) < 14:
                    rowc.append({'x0': fl['x0'], 'x1': fl['x1'],
                                 'top': fl['y'] - 1, 'bottom': fl['y'] + 1,
                                 'size': ref, 'font': 'MicrosoftYaHei',
                                 'text': '＿＿＿＿＿', 'page': page['page']})
            rowc.sort(key=lambda c: c['x0'])
            row_segs[k] = row_segments(rowc, ref, ctx)

        # 大括号分组：未配对左大括号纵向覆盖 >=2 行 -> 原生方程数组
        # 成员行：紧贴括号右侧；标签行（如"①恒成立问题："）：整体位于括号左侧
        groups = []
        consumed = set()
        for b in braces:
            inside = [k for k in range(len(rows))
                      if b['y0'] - 6 <= row_mid[k] <= b['y1'] + 6
                      and b['x'] - 2 <= min(c['x0'] for c in rows[k])
                      <= b['x'] + 40]
            labels = [k for k in range(len(rows))
                      if k not in inside
                      and b['y0'] - 6 <= row_mid[k] <= b['y1'] + 6
                      and max(c['x1'] for c in rows[k]) <= b['x'] + 2]
            if len(inside) >= 2:
                groups.append({'y': b['y0'], 'x': b['x'], 'ks': inside,
                               'label_ks': labels})
                consumed.update(inside)
                consumed.update(labels)

        for k, r in enumerate(rows):
            if k in consumed:
                continue
            items.append({'kind': 'row', 'y': row_mid[k],
                          'segs': row_segs[k], 'row': r, 'ref': ref})
        for g in groups:
            items.append({'kind': 'bracegroup', 'y': g['y'], 'g': g,
                          'segs_list': [row_segs[k] for k in g['ks']],
                          'label_segs': [row_segs[k]
                                         for k in g.get('label_ks', [])],
                          'rows': [rows[k] for k in g['ks']],
                          'ref': ref,
                          'msize': max(row_msize(rows[k], ref)
                                       for k in g['ks'])})

        for t in tbl_cells:
            items.append({'kind': 'table', 'y': t['y'], 't': t, 'ref': ref})

        items.sort(key=lambda it: (it['y'], 0 if it['kind'] == 'row' else 1))

        for it in items:
            if it['kind'] == 'table':
                t = it['t']
                tb = doc.add_table(rows=t['nrow'], cols=t['ncol'])
                tb.style = 'Table Grid'
                tb.alignment = WD_TABLE_ALIGNMENT.LEFT
                for i in range(t['nrow']):
                    for j in range(t['ncol']):
                        cell = tb.cell(i, j)
                        sub_rows = cluster_rows(t['grid'][i][j], ref)
                        for si, sr in enumerate(sub_rows):
                            s = row_segments(sr, ref, ctx)
                            if si == 0:
                                p = cell.paragraphs[0]
                            else:
                                p = cell.add_paragraph()
                            write_para(p, s)
                # 表格间隔段：保留（防止相邻表格粘连）但压到 3pt 高，
                # 免得空段把内容挤出尴尬的分页
                sp = doc.add_paragraph()
                sp.paragraph_format.space_before = Pt(0)
                sp.paragraph_format.space_after = Pt(0)
                ppr = sp._p.get_or_add_pPr()
                rpr = ppr.find(qn('w:rPr'))
                if rpr is None:
                    from docx.oxml import OxmlElement
                    rpr = OxmlElement('w:rPr')
                    ppr.append(rpr)
                for tagn in ('w:sz', 'w:szCs'):
                    e = OxmlElement(tagn)
                    e.set(qn('w:val'), '6')        # 3pt
                    rpr.append(e)
            elif it['kind'] == 'bracegroup':
                indent = max(it['g']['x'] - LEFT_PT, 0)
                p = add_brace_group(doc, it['segs_list'], indent, it['msize'],
                                    label_segs=it.get('label_segs'))
                n = len(p._p.findall('.//' + M_NS + 'oMath'))
                total_math += n
            else:
                r = it['row']
                segs = it['segs']
                txt = ''.join(v for _, v in segs)
                x0 = min(c['x0'] for c in r)
                size = max(c['size'] for c in r)
                bold = sum('Bold' in c['font'] for c in r) > len(r) / 2
                lvl = min(4, max(0, int(round((x0 - left) / 17.0))))
                indent = lvl * 17.0
                tsize = row_tsize(r, ref)
                msize = row_msize(r, ref)
                # 标题判定只看正文（非数学）字符：定义行里的大号 ∀/∃ 显示符
                # 不应把整行撑成标题
                body_max = max([round(c['size'], 1) for c in r
                                if c['text'].strip() and not is_math_char(c)],
                               default=0)

                align = None
                if abs((x0 + max(c['x1'] for c in r)) / 2 - page['width'] / 2) < 18 \
                        and size >= 11:
                    align = WD_ALIGN_PARAGRAPH.CENTER

                style = None
                if body_max >= 13 or txt.startswith('§') or HEAD_RE.match(txt) \
                        or re.match(r'^[一二三四五六七八九十]+、', txt):
                    style = 'Heading 1' if (body_max >= 15
                                            or txt.startswith('§')) \
                        else 'Heading 2'
                    indent = 0
                    if body_max:
                        tsize = body_max
                    align = WD_ALIGN_PARAGRAPH.CENTER if body_max >= 15 else None
                elif KEY_RE.match(txt):
                    bold = True

                add_rich(doc, segs, style=style, align=align,
                         indent_pt=indent, bold=bold,
                         tsize=tsize, msize=msize)

        if dump:
            print('--- PAGE %d ---' % page['page'])
            for it in items:
                if it['kind'] == 'row':
                    s = ''.join('$%s$' % v if k == 'math' else v
                                for k, v in it['segs'])
                    print('  %7.1f | %s' % (it['y'], s))
                elif it['kind'] == 'bracegroup':
                    print('  %7.1f | <BRACE-GROUP x=%.1f>'
                          % (it['y'], it['g']['x']))
                    for segs in it.get('label_segs', []):
                        s = ''.join(v for k, v in segs if k == 'text')
                        if s.strip():
                            print('          | %s {...' % s.strip())
                    for segs in it['segs_list']:
                        s = ''.join('$%s$' % v if k == 'math' else v
                                    for k, v in segs)
                        print('          |   { %s' % s)
                else:
                    print('  %7.1f | <TABLE %dx%d>'
                          % (it['y'], it['t']['nrow'], it['t']['ncol']))
                    for i in range(it['t']['nrow']):
                        cells = []
                        for j in range(it['t']['ncol']):
                            srs = cluster_rows(it['t']['grid'][i][j], ref)
                            s = ' / '.join(
                                ''.join('$%s$' % v if kk == 'math' else v
                                        for kk, v in row_segments(sr, ref, ctx))
                                for sr in srs)
                            cells.append(s)
                        print('           | ' + ' │ '.join(cells))

        # 不插硬分页：PDF 分页位置 ≠ Word 自然分页位置，
        # 硬分页会把内容拦腰截断留下半页空白，交给 Word 自动分页

    M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
    n_math = sum(1 for _ in doc.element.body.iter(M + 'oMath'))
    n_para = len(doc.paragraphs)
    n_para_math = sum(1 for _ in doc.element.body.iter(M + 'oMathPara'))
    try:
        doc.save(out_path)
    except PermissionError:
        root, ext = os.path.splitext(out_path)
        out_path = root + '_new' + ext
        doc.save(out_path)
        sys.stderr.write('[warn] 目标文件被占用（可能在 Word 中打开），'
                         '已另存为 %s\n' % out_path)
    print('  段落 %d，行内公式 %d，独立公式 %d' % (n_para, n_math, n_para_math))
    return out_path, n_math, failed


# ---------------------------------------------------------------- CLI
if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('pdf')
    ap.add_argument('-o', '--out', default=None)
    ap.add_argument('--dump', action='store_true')
    a = ap.parse_args()
    out = a.out or os.path.splitext(a.pdf)[0] + '.docx'
    p, nmath, nfail = convert(a.pdf, out, dump=a.dump)
    print('已生成: %s' % p)
    print('原生公式对象: %d，失败: %d' % (nmath, nfail))
