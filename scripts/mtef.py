# -*- coding: utf-8 -*-
"""
mtef — MathType MTEF v5 二进制格式解析器

背景：老讲义里的公式大多是 MathType OLE 对象。它在 docx 里存两份东西：

    word/media/*.wmf       呈现图（矢量，文字原样但**按字体分组**，无结构）
    word/embeddings/*.bin  OLE 复合文档 → "Equation Native" 流 → **MTEF 数据**

MTEF 里才有真正的公式结构（分数、根号、上下标、矩阵……）。
本模块把 MTEF 解析成对象树，再翻译成 LaTeX，接上 latex-omml 就能变成
Word 原生可编辑公式 —— 整条链路不需要 OCR、不需要装 MathType、不需要装 Office。

用法：
    python mtef.py 讲义.docx                    # 列出所有公式的 LaTeX
    python mtef.py 讲义.docx --json out.json    # 导出
    python mtef.py --selftest                   # 跑官方示例流的单元测试

=========================== 格式要点（全是踩过的坑）===========================

1. Equation Native 流 = 28 字节 OLE 头 + 12 字节 MathType 版本头 + MTEF 记录。
   （MEE / 公式编辑器 3.0 的版本头只有 5 字节，别搞混）
   版本头第 1 字节是 MTEF 版本号，必须是 5。

2. 记录的 options 是**独立字节**（type 后面）。MTEF v4 及更早版本把选项
   放在 type 字节的高 4 位 —— 如果你在解析 v5 数据却按老办法做，全盘皆错。

3. **不是所有记录都有 options 字节**。已确认没有的：
   END(0)、FONT_STYLE_DEF(8)、SIZE(9)、typesize 简写(10-14)、
   FONT_DEF(17)、ENCODING_DEF(19)、RULER(7)。
   有 options 的：LINE(1)、CHAR(2)、TMPL(3)、PILE(4)、MATRIX(5)、
   EMBELL(6)、EQN_PREFS(18)、COLOR_DEF(16)。

4. **EQN_PREFS(18) 不以 END 结尾**，长度由内部三个数组的 count 决定：
   type + options + sizes(dim array) + spaces(dim array) + styles(array)。
   dim array = count 字节 + nibble 流，每个 dimension 是
   [units nibble][十进制数字 nibble...][0xF 终止]。
   这是最容易把流读崩的地方。

5. NUDGE(0x08) 的 6 字节大偏移以 `80 80` 打头，2 字节小偏移是 `dx+128, dy+128`。

6. CHAR 的字符数据顺序固定：**MTCode(2 字节 LE) → 8-bit 字体码(0x04) →
   16-bit 字体码(0x10)**。MTCode 本身就是 Unicode 码位，优先用它。

7. typesize 简写记录在对象列表里会**穿插出现**（如 SIZE_SUB 出现在 TMPL 的
   两个 slot 之间），它们不影响结构，解析时必须跳过而不是当成 slot。

8. ScrBoxClass(tmSUB/tmSUP/tmSUBSUP) 的 slot 顺序恒为 **[下标, 上标]**。
   tmSUP 的下标 slot 是一个 options=0x01(LINE_NULL) 的空 LINE，
   LINE_NULL 的 LINE **连 END 都没有**，不要去读它的对象列表。

9. RootBoxClass(tmROOT) 的 slot 顺序是 **[被开方数, 根指数]**，
   平方根的根指数 slot 是空 LINE。

10. BigOpBoxClass(tmINTEG/tmSUM/...) 的 slot 是
    **[主体, 上限, 下限, 算子字符]**，最后那个是 CHAR 不是 LINE。
    注意主体在最前、算子在最后 —— 和 OMML 的顺序相反。
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


# ------------------------------------------------------------------ 常量表

END, LINE, CHAR, TMPL, PILE, MATRIX, EMBELL, RULER = range(8)
FONT_STYLE_DEF, SIZE, FULL, SUB, SUB2, SYM, SUBSYM = 8, 9, 10, 11, 12, 13, 14
COLOR, COLOR_DEF, FONT_DEF, EQN_PREFS, ENCODING_DEF = 15, 16, 17, 18, 19

REC_NAME = {0: 'END', 1: 'LINE', 2: 'CHAR', 3: 'TMPL', 4: 'PILE', 5: 'MATRIX',
            6: 'EMBELL', 7: 'RULER', 8: 'FONT_STYLE_DEF', 9: 'SIZE', 10: 'FULL',
            11: 'SUB', 12: 'SUB2', 13: 'SYM', 14: 'SUBSYM', 15: 'COLOR',
            16: 'COLOR_DEF', 17: 'FONT_DEF', 18: 'EQN_PREFS', 19: 'ENCODING_DEF'}

# 不带 options 字节的记录类型（见文件头说明 3）
NO_OPTION = {END, FONT_STYLE_DEF, SIZE, FULL, SUB, SUB2, SYM, SUBSYM,
             FONT_DEF, ENCODING_DEF, RULER, COLOR}

OPT_NUDGE = 0x08
OPT_CHAR_EMBELL = 0x01
OPT_CHAR_ENC_CHAR_8 = 0x04
OPT_CHAR_ENC_CHAR_16 = 0x10
OPT_CHAR_ENC_NO_MTCODE = 0x20
OPT_LINE_NULL = 0x01
OPT_LP_RULER = 0x02
OPT_LINE_LSPACE = 0x04

# MathType 逻辑字体（CHAR 的 typeface 是「逻辑值 + 128」后按 signed integer 存）
TYPEFACE = {1: 'text', 2: 'function', 3: 'variable', 4: 'lcgreek', 5: 'ucgreek',
            6: 'symbol', 7: 'vector', 8: 'number', 9: 'user1', 10: 'user2',
            11: 'mtextra', 12: 'text_fe', 22: 'expand', 23: 'marker', 24: 'space'}

# TMPL selector
(tmANGLE, tmPAREN, tmBRACE, tmBRACK, tmBAR, tmDBAR, tmFLOOR, tmCEILING,
 tmOBRACK) = range(9)
tmINTERVAL, tmROOT, tmFRACT, tmUBAR, tmOBAR, tmARROW, tmINTEG = 9, 10, 11, 12, 13, 14, 15
tmSUM, tmPROD, tmCOPROD, tmUNION, tmINTER, tmINTOP, tmSUMOP = 16, 17, 18, 19, 20, 21, 22
tmLIM, tmHBRACE, tmHBRACK, tmLDIV = 23, 24, 25, 26
tmSUB, tmSUP, tmSUBSUP, tmDIRAC = 27, 28, 29, 30
tmVEC, tmTILDE, tmHAT, tmARC, tmJSTATUS, tmSTRIKE, tmBOX = 31, 32, 33, 34, 35, 36, 37

SEL_NAME = {0: 'ANGLE', 1: 'PAREN', 2: 'BRACE', 3: 'BRACK', 4: 'BAR', 5: 'DBAR',
            6: 'FLOOR', 7: 'CEILING', 8: 'OBRACK', 9: 'INTERVAL', 10: 'ROOT',
            11: 'FRACT', 12: 'UBAR', 13: 'OBAR', 14: 'ARROW', 15: 'INTEG',
            16: 'SUM', 17: 'PROD', 18: 'COPROD', 19: 'UNION', 20: 'INTER',
            21: 'INTOP', 22: 'SUMOP', 23: 'LIM', 24: 'HBRACE', 25: 'HBRACK',
            26: 'LDIV', 27: 'SUB', 28: 'SUP', 29: 'SUBSUP', 30: 'DIRAC',
            31: 'VEC', 32: 'TILDE', 33: 'HAT', 34: 'ARC', 35: 'JSTATUS',
            36: 'STRIKE', 37: 'BOX'}

# fence 的开/闭符号（selector → (left, right)）
FENCE_PAIR = {
    tmANGLE:   ('\\langle ', '\\rangle '),
    tmPAREN:   ('(', ')'),
    tmBRACE:   ('\\{', '\\}'),
    tmBRACK:   ('[', ']'),
    tmBAR:     ('|', '|'),
    tmDBAR:    ('\\|', '\\|'),
    tmFLOOR:   ('\\lfloor ', '\\rfloor '),
    tmCEILING: ('\\lceil ', '\\rceil '),
    tmOBRACK:  ('\\llbracket ', '\\rrbracket '),
    tmINTERVAL: ('(', ')'),        # 区间默认半开，实际字符以 slot 里存的为准
}

# slot 里存得下的 fence 字符白名单。
# 只认这些才把 slot1/slot2 当括号用 —— 防止把别的内容误当成 fence。
FENCE_CHARS = set('()[]{}<>|‖⌊⌋⌈⌉⟦⟧⟨⟩/\\ ⟦⟧') | {
    '\\langle ', '\\rangle ', '\\lfloor ', '\\rfloor ',
    '\\lceil ', '\\rceil ', '\\llbracket ', '\\rrbracket ',
    '\\{', '\\}', '\\|',
}

# 大运算符（BigOpBoxClass / LimBoxClass）的 LaTeX
BIGOP = {tmINTEG: '\\int ', tmSUM: '\\sum ', tmPROD: '\\prod ',
         tmCOPROD: '\\coprod ', tmUNION: '\\bigcup ', tmINTER: '\\bigcap ',
         tmINTOP: '\\int ', tmSUMOP: '\\sum ', tmLIM: '\\lim '}

# EMBELL 装饰类型
EMBELL_LATEX = {
    2: ('\\dot{', '}'),      3: ('\\ddot{', '}'),      4: ('\\dddot{', '}'),
    5: ('{', "'}" ),         6: ('{', "''}"),          7: ("{}^{\\backprime}", ''),
    8: ('\\tilde{', '}'),    9: ('\\hat{', '}'),      10: ('\\not ', ''),
    11: ('\\vec{', '}'),     12: ('\\overset{\\leftarrow}{', '}'),
    13: ('\\overset{\\leftrightarrow}{', '}'),
    16: ('\\overline{', '}'),
    17: ('\\overline{', '}'),
    18: ('{', "'''}"),
    19: ('\\overset{\\frown}{', '}'),
    20: ('\\overset{\\smile}{', '}'),
    24: ('\\ddddot{', '}'),
    29: ('\\underline{', '}'),
    30: ('\\underset{\\sim}{', '}'),
    31: ('\\overset{\\frown}{', '}'),
    32: ('\\overset{\\smile}{', '}'),
    33: ('\\underset{\\rightarrow}{', '}'),
}

# LaTeX 需要转义的字符
LATEX_SPECIAL = {'\\': '\\backslash ', '{': '\\{', '}': '\\}', '$': '\\$',
                 '%': '\\%', '&': '\\&', '#': '\\#', '_': '\\_', '~': '\\sim '}

# MTCode（本身就是 Unicode 码位）→ LaTeX 命令
# 不转也能用，但 latex2mathml 对 LaTeX 命令的支持比对裸 Unicode 稳得多
# （\perp 的码位 bug 就是活生生的教训）
MTCODE_LATEX = {
    0x00B1: '\\pm ',     0x00D7: '\\times ',   0x00F7: '\\div ',
    0x2212: '-',         0x2260: '\\neq ',     0x2248: '\\approx ',
    0x2261: '\\equiv ',  0x2264: '\\leq ',     0x2265: '\\geq ',
    0x226A: '\\ll ',     0x226B: '\\gg ',
    0x2208: '\\in ',     0x2209: '\\notin ',   0x220B: '\\ni ',
    0x2282: '\\subset ', 0x2283: '\\supset ',  0x2286: '\\subseteq ',
    0x2287: '\\supseteq ', 0x2ABB: '\\subsetneqq ',
    0x2229: '\\cap ',    0x222A: '\\cup ',     0x2205: '\\varnothing ',
    0x221E: '\\infty ',  0x221A: '\\sqrt ',    0x2202: '\\partial ',
    0x2207: '\\nabla ',  0x2220: '\\angle ',   0x22A5: '\\perp ',
    0x2225: '\\parallel ', 0x22C5: '\\cdot ',  0x2219: '\\bullet ',
    0x00B0: '^{\\circ}', 0x2032: "'",          0x2033: "''",
    0x2192: '\\to ',     0x2190: '\\leftarrow ', 0x2194: '\\leftrightarrow ',
    0x21D2: '\\Rightarrow ', 0x21D0: '\\Leftarrow ',
    0x21D4: '\\Leftrightarrow ', 0x2200: '\\forall ', 0x2203: '\\exists ',
    0x2234: '\\therefore ', 0x2235: '\\because ',
    0x2211: '\\sum ',    0x222B: '\\int ',     0x220F: '\\prod ',
    0x2245: '\\cong ',   0x223C: '\\sim ',
    0x2329: '\\langle ', 0x232A: '\\rangle ',
    0x3008: '\\langle ', 0x3009: '\\rangle ',
    0x00A0: '\\ ',       0x2009: '\\,',        0x2002: '\\quad ',
    # MathType 私有区（fnEXPAND / typeface=22）：可伸缩括号的**部件**
    # 这些不是语义字符，只是排版用的拉伸件，统一映射成对应的普通符号
    0xEC07: '|',   0xEC08: '|',        # 可伸缩竖线的左/右半
    0xEC09: '\\|', 0xEC0A: '\\|',      # 可伸缩双竖线
    0xEF04: '|',   0xEF08: '\\|',      # MTExtra 单/双竖线
    0xEC0B: '(',   0xEC0C: ')',
    0xEC0D: '[',   0xEC0E: ']',
    0xEC0F: '\\{', 0xEC10: '\\}',
}

# 私有区里没映射到的码位：都是排版部件，不是语义字符，直接丢弃
_MT_PRIVATE = range(0xE000, 0xF900)
_GREEK = {
    'alpha': 0x03B1, 'beta': 0x03B2, 'gamma': 0x03B3, 'delta': 0x03B4,
    'epsilon': 0x03B5, 'zeta': 0x03B6, 'eta': 0x03B7, 'theta': 0x03B8,
    'iota': 0x03B9, 'kappa': 0x03BA, 'lambda': 0x03BB, 'mu': 0x03BC,
    'nu': 0x03BD, 'xi': 0x03BE, 'omicron': 0x03BF, 'pi': 0x03C0,
    'rho': 0x03C1, 'sigmaf': 0x03C2, 'sigma': 0x03C3, 'tau': 0x03C4,
    'upsilon': 0x03C5, 'phi': 0x03C6, 'chi': 0x03C7, 'psi': 0x03C8,
    'omega': 0x03C9,
}
for _n, _c in _GREEK.items():
    MTCODE_LATEX[_c] = '\\%s ' % _n
MTCODE_LATEX.update({
    0x0391: '\\Alpha ', 0x0392: '\\Beta ', 0x0393: '\\Gamma ',
    0x0394: '\\Delta ', 0x0395: '\\Epsilon ', 0x0396: '\\Zeta ',
    0x0397: '\\Eta ', 0x0398: '\\Theta ', 0x0399: '\\Iota ',
    0x039A: '\\Kappa ', 0x039B: '\\Lambda ', 0x039C: '\\Mu ',
    0x039D: '\\Nu ', 0x039E: '\\Xi ', 0x039F: '\\Omicron ',
    0x03A0: '\\Pi ', 0x03A1: '\\Rho ', 0x03A3: '\\Sigma ',
    0x03A4: '\\Tau ', 0x03A5: '\\Upsilon ', 0x03A6: '\\Phi ',
    0x03A7: '\\Chi ', 0x03A8: '\\Psi ', 0x03A9: '\\Omega ',
})


# ------------------------------------------------------------------ 字节读取

class MTEFError(Exception):
    pass


class Reader(object):
    """带边界检查的字节流读取器"""

    def __init__(self, data, pos=0):
        self.d = data
        self.i = pos

    def eof(self):
        return self.i >= len(self.d)

    def u8(self):
        if self.i >= len(self.d):
            raise MTEFError('读越界 at %d' % self.i)
        b = self.d[self.i]
        self.i += 1
        return b

    def peek(self):
        if self.i >= len(self.d):
            raise MTEFError('peek 越界 at %d' % self.i)
        return self.d[self.i]

    def raw(self, n):
        if self.i + n > len(self.d):
            raise MTEFError('读 %d 字节越界 at %d' % (n, self.i))
        b = self.d[self.i:self.i + n]
        self.i += n
        return b

    def u16(self):
        """simple 16-bit：固定 2 字节，低字节在前"""
        lo, hi = self.raw(2)
        return lo | (hi << 8)

    def signed_int(self):
        """signed integer：1 字节(value+128)，或 FF + 2 字节(value+32768)"""
        b = self.u8()
        if b != 0xFF:
            return b - 128
        return self.u16() - 32768

    def unsigned_int(self):
        """unsigned integer：1 字节，或 FF + 2 字节"""
        b = self.u8()
        if b != 0xFF:
            return b
        return self.u16()

    def nudge(self):
        if self.peek() == 0x80 and self.d[self.i + 1:self.i + 2] == b'\x80':
            self.i += 2
            dx = struct.unpack('<h', self.raw(2))[0]
            dy = struct.unpack('<h', self.raw(2))[0]
            return (dx, dy)
        dx = self.u8() - 128
        dy = self.u8() - 128
        return (dx, dy)

    def cstring(self):
        end = self.d.find(b'\x00', self.i)
        if end < 0:
            raise MTEFError('cstring 没有终止符')
        s = self.d[self.i:end]
        self.i = end + 1
        return s


class NibbleStream(object):
    """nibble 流：每个字节拆成两个 4 bit，高 4 位先出"""

    def __init__(self, r):
        self.r = r
        self.pending = None

    def next(self):
        if self.pending is not None:
            n = self.pending
            self.pending = None
            return n
        b = self.r.u8()
        self.pending = b & 0x0F
        return (b >> 4) & 0x0F

    def align(self):
        """丢弃半个字节的残留，回到字节边界"""
        self.pending = None


# ------------------------------------------------------------------ 对象树

class Node(object):
    def latex(self):
        raise NotImplementedError


class Text(Node):
    """一个字符（或一串纯字符被合并后的片段）"""

    def __init__(self, s, typeface=None, embell=None, mtcode=None):
        self.s = s
        self.typeface = typeface
        self.embell = embell or []
        self.mtcode = mtcode

    def latex(self):
        if self.mtcode is not None and self.mtcode in MTCODE_LATEX:
            s = MTCODE_LATEX[self.mtcode]
        elif self.mtcode in _MT_PRIVATE:
            s = ''          # 未登记的私有区码位 = 排版部件，丢掉
        else:
            s = ''.join(LATEX_SPECIAL.get(c, c) for c in self.s)
        for e in self.embell:
            pre, post = EMBELL_LATEX.get(e, ('', ''))
            s = pre + s + post
        return s


class Line(Node):
    def __init__(self, items, null=False):
        self.items = items
        self.null = null

    def latex(self):
        return ''.join(i.latex() for i in self.items)


class Tmpl(Node):
    def __init__(self, selector, variation, options, slots):
        self.selector = selector
        self.variation = variation
        self.options = options
        self.slots = slots      # list[Line | Text]

    def latex(self):
        return tmpl_latex(self)


class Pile(Node):
    def __init__(self, lines, halign, valign):
        self.lines = lines
        self.halign = halign
        self.valign = valign

    def latex(self):
        return node_latex(self)


class Mtx(Node):
    def __init__(self, rows, cols, cells):
        self.rows = rows
        self.cols = cols
        self.cells = cells

    def latex(self):
        return node_latex(self)


# ------------------------------------------------------------------ 解析

class MTEFParser(object):
    def __init__(self, body):
        self.r = Reader(body)
        self.font_defs = []      # FONT_DEF，索引从 1 开始
        self.styles = []         # EQN_PREFS 里的 style 定义
        self.tex_input = None    # FUTURE 0x66 里夹带的原始 TeX 源码（如果有）

    # ---- FUTURE 记录里的附加信息 ----
    def _note_future(self, t, blob):
        """FUTURE 记录夹带的附加信息

        目前只认一种：type 0x66 = "TeX Input Language"，内容为
            b'TeX Input Language\\0' + <作者当初敲的原始 LaTeX> + b'\\0'
        这份原文比从 MTEF 结构翻译回来的结果**更可靠**（没有启发式猜测）。
        MathType 从 6.x 起，凡是用 TeX 输入方式建的公式都会带。
        """
        if t != 0x66:
            return
        i = blob.find(b'\x00')
        if i < 0:
            return
        name, rest = blob[:i], blob[i + 1:]
        if name != b'TeX Input Language':
            return
        s = rest.split(b'\x00')[0]
        try:
            self.tex_input = s.decode('utf-8')
        except UnicodeDecodeError:      # 老版本可能用 GBK / latin-1
            for enc in ('gbk', 'latin-1'):
                try:
                    self.tex_input = s.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue

    # ---- 顶层入口 ----
    def parse(self):
        while not self.r.eof():
            rec = self.read_record(top=True)
            if rec is None:
                continue
            if isinstance(rec, (Line, Pile)):
                return rec
        return None

    # ---- 记录分发 ----
    def read_record(self, top=False):
        t = self.r.u8()

        # FUTURE(>=100) 必须在这里就分流，不能走下面的通用 options 逻辑：
        # 它的格式是 type + 长度 + 数据，**根本没有 options 字节**。
        # 把长度字节当成 options 读，除整体错位 1 字节外，
        # 长度字节只要 bit3 为 1（如 0x1e）就会被误判成 OPT_NUDGE 再吞 2 字节，
        # 于是从这条记录起后面全废（实测一份 265 字节的流直接报「未知类型 0x43」）。
        #
        # 顺带：MathType 用 FUTURE 0x66 存「TeX Input Language」——
        # 公式的**原始 LaTeX 源码**（形如 b'TeX Input Language\0<a>\0'），
        # 比从 MTEF 结构翻译回来更准。见 self.tex_input。
        if t >= 100:
            n = self.r.unsigned_int()
            blob = self.r.raw(n)
            self._note_future(t, blob)
            return None

        opt = 0
        if t not in NO_OPTION:
            opt = self.r.u8()

        if t == END:
            return None
        if t in (FULL, SUB, SUB2, SYM, SUBSYM):
            return None                      # typesize 简写，跳过
        if t == SIZE:
            self.skip_size()
            return None
        if t == COLOR:
            self.r.unsigned_int()
            return None
        if t == FONT_STYLE_DEF:
            self.r.unsigned_int()
            self.r.u8()
            return None
        if t == ENCODING_DEF:
            self.r.cstring()
            return None
        if t == FONT_DEF:
            enc = self.r.unsigned_int()
            name = self.r.cstring()
            self.font_defs.append((enc, name))
            return None
        if t == EQN_PREFS:
            self.read_eqn_prefs()
            return None
        if t == COLOR_DEF:
            self.skip_color_def(opt)
            return None
        if t == RULER:
            self.skip_ruler()
            return None

        if opt & OPT_NUDGE:
            self.r.nudge()

        if t == LINE:
            return self.read_line(opt)
        if t == CHAR:
            return self.read_char(opt)
        if t == TMPL:
            return self.read_tmpl()
        if t == PILE:
            return self.read_pile(opt)
        if t == MATRIX:
            return self.read_matrix()
        if t == EMBELL:
            self.r.u8()
            return None


        raise MTEFError('未知记录类型 %d at %d' % (t, self.r.i))

    # ---- 各类记录 ----
    def read_line(self, opt):
        if opt & OPT_LINE_NULL:
            return Line([], null=True)        # 连 END 都没有
        if opt & OPT_LINE_LSPACE:
            self.r.u16()
        if opt & OPT_LP_RULER:
            self.skip_ruler()
        return Line(self.read_object_list())

    def read_char(self, opt):
        tf = self.r.signed_int()
        mtcode = None
        fontpos = None
        if not (opt & OPT_CHAR_ENC_NO_MTCODE):
            mtcode = self.r.u16()
        if opt & OPT_CHAR_ENC_CHAR_8:
            fontpos = self.r.u8()
        elif opt & OPT_CHAR_ENC_CHAR_16:
            fontpos = self.r.u16()

        # 字符取值：MTCode 本身就是 Unicode 码位，优先
        if mtcode is not None and mtcode:
            ch = chr(mtcode)
        elif fontpos is not None:
            ch = chr(fontpos)
        else:
            ch = ''

        embell = []
        if opt & OPT_CHAR_EMBELL:
            while True:
                if self.r.peek() == END:
                    self.r.u8()
                    break
                t = self.r.u8()
                if t not in NO_OPTION:
                    o = self.r.u8()
                    if o & OPT_NUDGE:
                        self.r.nudge()
                embell.append(self.r.u8())

        return Text(ch, typeface=tf, embell=embell, mtcode=mtcode)

    def read_tmpl(self):
        sel = self.r.u8()
        v1 = self.r.u8()
        if v1 & 0x80:
            v2 = self.r.u8()
            var = (v1 & 0x7F) | (v2 << 8)
        else:
            var = v1
        self.r.u8()                            # template-specific options
        slots = self.read_object_list()
        return Tmpl(sel, var, 0, slots)

    def read_pile(self, opt):
        halign = self.r.u8()
        valign = self.r.u8()
        if opt & OPT_LP_RULER:
            self.skip_ruler()
        lines = self.read_object_list()
        return Pile(lines, halign, valign)

    def read_matrix(self):
        valign = self.r.u8()
        hjust = self.r.u8()
        vjust = self.r.u8()
        rows = self.r.u8()
        cols = self.r.u8()
        # row_parts / col_parts：每个分隔线 2 bit，数量 rows+1 / cols+1
        for n in (rows + 1, cols + 1):
            self.r.raw((n * 2 + 7) // 8)
        cells = self.read_object_list()
        return Mtx(rows, cols, cells)

    def read_object_list(self):
        """读到 END 为止，跳过所有非结构性记录"""
        items = []
        while True:
            if self.r.eof():
                raise MTEFError('对象列表未闭合')
            if self.r.peek() == END:
                self.r.u8()
                break
            node = self.read_record()
            if node is not None:
                items.append(node)
        return items

    # ---- EQN_PREFS（最容易读崩的地方）----
    def read_eqn_prefs(self):
        self.read_dim_array()
        self.read_dim_array()
        n = self.r.u8()
        self.styles = []
        for _ in range(n):
            idx = self.r.unsigned_int()
            if idx:
                self.r.u8()                    # character style
            self.styles.append(idx)

    def read_dim_array(self):
        count = self.r.u8()
        ns = NibbleStream(self.r)
        for _ in range(count):
            ns.next()                          # units nibble
            while ns.next() != 0x0F:           # 数字 nibble，直到终止符
                pass
        ns.align()

    # ---- 跳过类 ----
    def skip_size(self):
        b = self.r.peek()
        if b == 100:                           # 大 delta
            self.r.u8()
            self.r.u8()
            self.r.u16()
        elif b == 101:                         # 显式点大小
            self.r.u8()
            self.r.u16()
        else:                                  # 小 delta
            self.r.u8()
            self.r.u8()

    def skip_color_def(self, opt):
        n = 4 if (opt & 0x01) else 3
        for _ in range(n):
            self.r.u16()
        if opt & 0x04:
            self.r.cstring()

    def skip_ruler(self):
        n = self.r.u8()
        for _ in range(n):
            self.r.u8()
            self.r.u16()


# ------------------------------------------------------------------ → LaTeX

def _slot_latex(slot, brace=True):
    """slot 可能是 Line / Text / Tmpl / None"""
    if slot is None:
        return ''
    s = slot.latex()
    if not s:
        return ''
    if brace and not (len(s) == 1 and s.isalnum()):
        return '{' + s + '}'
    return s


def node_latex(node):
    if isinstance(node, Text):
        return node.latex()      # 走 MTCode→LaTeX 映射，别自己拼字符

    if isinstance(node, Line):
        return ''.join(node_latex(i) for i in node.items)

    if isinstance(node, Pile):
        # 多行堆叠。中小学讲义里绝大多数是分段函数 / 方程组
        rows = [node_latex(r) for r in node.lines]
        rows = [r for r in rows if r != '']
        if not rows:
            return ''
        if len(rows) == 1:
            return rows[0]
        # 含 & 的是对齐结构（cases），否则用 matrix
        if any('&' in r for r in rows):
            return '\\begin{cases}' + ' \\\\ '.join(rows) + '\\end{cases}'
        return '\\begin{matrix}' + ' \\\\ '.join(rows) + '\\end{matrix}'

    if isinstance(node, Mtx):
        rows = []
        for r in range(node.rows):
            cells = []
            for c in range(node.cols):
                idx = r * node.cols + c
                cells.append(node_latex(node.cells[idx]) if idx < len(node.cells) else '')
            rows.append(' & '.join(cells))
        return '\\begin{matrix}' + ' \\\\ '.join(rows) + '\\end{matrix}'

    if isinstance(node, Tmpl):
        return tmpl_latex(node)

    return ''


def _slots(t, n):
    """取前 n 个 slot，不足补 None"""
    out = list(t.slots[:n])
    while len(out) < n:
        out.append(None)
    return out


def _is_empty(slot):
    if slot is None:
        return True
    if isinstance(slot, Line) and slot.null:
        return True
    return slot.latex() == ''


def tmpl_latex(t):
    sel, var = t.selector, t.variation

    # ---- 括号类（含 INTERVAL=9，别写成 sel <= tmOBRACK，会漏掉区间）----
    if sel <= tmINTERVAL:
        body = _slot_latex(_slots(t, 3)[0], brace=False)
        lf, rf = FENCE_PAIR.get(sel, ('(', ')'))
        sl = _slots(t, 3)
        # 开闭符号直接读 slot1/slot2 里存的实际字符，**不要**看 variation 位。
        # 原因：INTERVAL 的 variation 编码（实测 0x30 / 0x12）跟 BRACK（0x3）
        # 不是一套位定义，按 var 位判断会漏掉区间符号，
        # 结果把 "(" 和 "]" 当成普通内容拼到末尾，输出成 `-∞,3(]` 这种。
        custom = False
        for idx, side in ((1, 0), (2, 1)):
            if _is_empty(sl[idx]):
                continue
            ch = sl[idx].latex()
            if ch.strip() in FENCE_CHARS or ch in FENCE_CHARS:
                if side == 0:
                    lf = ch
                else:
                    rf = ch
                custom = True
        # 自定义字符跟着内容高度伸缩，必须用 \left...\right
        if custom:
            return '\\left' + lf + body + '\\right' + rf
        return lf + body + rf

    # ---- 根式：[被开方数, 根指数] ----
    if sel == tmROOT:
        rad, idx = _slots(t, 2)
        if _is_empty(idx):
            return '\\sqrt{' + node_latex(rad) + '}'
        return '\\sqrt[' + node_latex(idx) + ']{' + node_latex(rad) + '}'

    # ---- 分式：[分子, 分母] ----
    if sel == tmFRACT:
        num, den = _slots(t, 2)
        if var & 0x0002:                       # 斜分数线
            return node_latex(num) + '/' + node_latex(den)
        return '\\frac{' + node_latex(num) + '}{' + node_latex(den) + '}'

    # ---- 上下划线 ----
    if sel == tmUBAR:
        return '\\underline{' + node_latex(_slots(t, 1)[0]) + '}'
    if sel == tmOBAR:
        return '\\overline{' + node_latex(_slots(t, 1)[0]) + '}'

    # ---- 上下标：[下标, 上标]（注意顺序！）----
    if sel in (tmSUB, tmSUP, tmSUBSUP):
        sub, sup = _slots(t, 2)
        base = ''
        # 前导脚本（tvSU_PRECEDES）时主体还没有，交给调用方拼接，这里只给脚本
        out = ''
        if not _is_empty(sub):
            out += '_' + _slot_latex(sub)
        if not _is_empty(sup):
            out += '^' + _slot_latex(sup)
        return out

    # ---- 大运算符：[主体, 上限, 下限, 算子] ----
    if sel in BIGOP:
        main, upper, lower, op = _slots(t, 4)
        name = BIGOP[sel]
        if sel == tmINTEG:
            n = var & 0x000F
            name = {1: '\\int ', 2: '\\iint ', 3: '\\iiint '}.get(n, '\\int ')
        lim = ''
        if not _is_empty(lower):
            lim += '_' + _slot_latex(lower)
        if not _is_empty(upper):
            lim += '^' + _slot_latex(upper)
        return name + lim + ' ' + node_latex(main)

    # ---- 向量 / 帽子 / 波浪 / 弧 ----
    if sel == tmVEC:
        # 启发式修正：MathType 把「向量的模」存成 VEC(BAR(b))，
        # 但呈现图（WMF）里箭头只有 2 段、只盖住单个字符 —— 也就是说
        # 它实际渲染成的是 |vec(b)| 而不是 vec(|b|)。
        # 数学上「|向量b|」也才说得通，所以把箭头挪进 fence 里面。
        inner = _slots(t, 1)[0]
        if (isinstance(inner, Line) and len(inner.items) == 1
                and isinstance(inner.items[0], Tmpl)
                and inner.items[0].selector <= tmOBRACK):
            fence = inner.items[0]
            parts = _slots(fence, 3)
            if parts[0] is not None and not _is_empty(parts[0]):
                arrow = _slots(t, 2)[1]
                new_vec = Tmpl(sel, var, 0,
                               [parts[0]] + ([arrow] if arrow is not None else []))
                fence.slots = [new_vec, parts[1], parts[2]]
                return fence.latex()

        body = node_latex(_slots(t, 1)[0])
        left = bool(var & 0x0001)
        under = bool(var & 0x0004)
        if under:
            return '\\underset{\\rightarrow}{' + body + '}'
        if left:
            return '\\overset{\\leftarrow}{' + body + '}'
        if len(body) <= 1:
            return '\\vec{' + body + '}'
        return '\\overrightarrow{' + body + '}'
    if sel == tmTILDE:
        body = node_latex(_slots(t, 1)[0])
        return ('\\widetilde{' if len(body) > 1 else '\\tilde{') + body + '}'
    if sel == tmHAT:
        body = node_latex(_slots(t, 1)[0])
        return ('\\widehat{' if len(body) > 1 else '\\hat{') + body + '}'
    if sel == tmARC:
        return '\\overset{\\frown}{' + node_latex(_slots(t, 1)[0]) + '}'

    # ---- 水平花括号 / 方括号 ----
    if sel in (tmHBRACE, tmHBRACK):
        body = node_latex(_slots(t, 1)[0])
        return '\\underbrace{' + body + '}' if sel == tmHBRACE else body

    # ---- lim ----
    if sel == tmLIM:
        main, lower, upper = _slots(t, 3)
        lim = ''
        if not _is_empty(lower):
            lim += '_' + _slot_latex(lower)
        return '\\lim' + lim + ' ' + node_latex(main)

    # ---- 未特化的模板：退化成把各 slot 拼起来，至少不丢字符 ----
    parts = [node_latex(s) for s in t.slots if not _is_empty(s)]
    return ''.join(parts)


def to_latex(node):
    return node_latex(node) if node is not None else ''


# ------------------------------------------------------------------ 提取

# OLE2 / CFBF 最小只读实现 —— olefile 装了就用 olefile，没装走这里。
# 之所以要自己写一份：换机器时依赖可能装不全，而 MathType 公式是核心功能，
# 不能因为少一个第三方包就全瘫。
#
# 关键前提（实测）：MathType 的 Equation Native 流只有 200~500 字节，
# **一律走迷你流**（CFBF 里 <4096B 的流放 mini stream，靠 mini FAT 串起来）。
# 所以只实现普通 FAT 是不够的，mini FAT 那段必须有。
_CFBF_SIG = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
FREESECT = 0xFFFFFFF0          # >= 这个值表示链结束


class _CFBF:
    """只读 CFBF v3（512B 扇区），只做两件事：找流名、读流内容"""

    def __init__(self, data):
        if data[:8] != _CFBF_SIG:
            raise MTEFError('不是 OLE 复合文档（CFBF 签名不对）')
        self.d = data
        self.ssz = 1 << int.from_bytes(data[0x1E:0x20], 'little')    # 扇区大小
        self.msz = 1 << int.from_bytes(data[0x20:0x22], 'little')    # 迷你扇区大小
        self.cutoff = int.from_bytes(data[0x38:0x3C], 'little') or 4096
        self.fat = self._read_fat()
        # 注意：_read_chain 返回的是**字节串**，FAT/MINIFAT 必须 unpack 成 uint32
        # 数组才能当索引用。直接拿 bytes 当数组用会取到单字节（0~255），
        # 既永远小于 FREESECT 又指向错误扇区 —— 表现为 _read_mini 死循环。
        self.minifat = self._u32array(self._read_chain(data[0x3C:0x40], self.ssz))
        self.dir = self._read_chain(data[0x30:0x34], self.ssz)
        self._ministream = None

    @staticmethod
    def _u32array(raw):
        """字节串 → uint32 列表（长度不足 4 的尾巴丢掉）"""
        n = len(raw) // 4
        return list(struct.unpack('<%dI' % n, raw[:4 * n])) if n else []

    # -- 基础设施 ------------------------------------------------------
    def _sec(self, n):
        """第 n 个扇区的数据（v3：头本身占第 0 个扇区位置）"""
        off = (n + 1) * self.ssz
        return self.d[off:off + self.ssz]

    def _chain(self, start, arr):
        """按 FAT 数组追链，返回扇区号列表"""
        out, n, guard = [], start, len(arr)
        while n < FREESECT and guard > 0:
            out.append(n)
            n = arr[n]
            guard -= 1
        return out

    def _read_chain(self, start_field, size):
        start = int.from_bytes(start_field, 'little')
        if start >= FREESECT:
            return b''
        return b''.join(self._sec(s) for s in self._chain(start, self.fat))

    def _read_fat(self):
        """DIFAT 前 109 项写在头里，小文件足够（MathType 只有 1~2 个 FAT 扇区）"""
        fat = []
        for i in range(109):
            s = int.from_bytes(self.d[0x4C + 4 * i:0x50 + 4 * i], 'little')
            if s >= FREESECT:
                break
            fat.extend(struct.unpack('<%dI' % (self.ssz // 4), self._sec(s)))
        return fat

    # -- 目录 ----------------------------------------------------------
    def entries(self):
        """yield (name, start_sector, size, objtype)；objtype: 2=流 5=根"""
        d = self.dir
        for i in range(0, len(d) - 127, 128):
            e = d[i:i + 128]
            nlen = int.from_bytes(e[0x40:0x42], 'little')
            if nlen < 2 or e[0x42] not in (2, 5):
                continue
            yield (e[:nlen - 2].decode('utf-16-le', 'replace'),
                   int.from_bytes(e[0x74:0x78], 'little'),
                   int.from_bytes(e[0x78:0x80], 'little'),
                   e[0x42])

    # -- 读流 ----------------------------------------------------------
    def openstream(self, name):
        ents = list(self.entries())
        root = next((e for e in ents if e[3] == 5), None)
        for nm, st, sz, ty in ents:
            if nm != name or ty != 2:
                continue
            if sz < self.cutoff:
                if root is None:
                    raise MTEFError('OLE 目录里没有根入口，读不了迷你流')
                return self._read_mini(st, sz, root[1])
            return self._read_chain(st.to_bytes(4, 'little'), self.ssz)[:sz]
        raise MTEFError('OLE 里没有流 %r' % name)

    def _read_mini(self, start, size, root_start):
        if self._ministream is None:
            self._ministream = self._read_chain(
                root_start.to_bytes(4, 'little'), self.ssz)
        out, n = b'', start
        # guard 是防御性的：mini FAT 若有环、或链索引越界，宁可截断也别死循环
        guard = len(self.minifat) + 1
        while len(out) < size and n < FREESECT and guard > 0:
            off = n * self.msz
            chunk = self._ministream[off:off + self.msz]
            if not chunk:
                break                      # 索引越界，链断了
            out += chunk
            n = self.minifat[n] if n < len(self.minifat) else FREESECT
            guard -= 1
        return out[:size]


def _ole_stream(data, name):
    """OLE 复合文档 → 指定流的字节

    优先用 olefile（成熟库，覆盖各种边界情况）；没装就回退到内置 _CFBF。
    实测两种路径对 MathType 文档结果完全一致。
    """
    try:
        import io as _io
        import olefile
    except ImportError:
        return _CFBF(data).openstream(name)
    ole = olefile.OleFileIO(_io.BytesIO(data))
    try:
        if not ole.exists(name):
            raise MTEFError('OLE 里没有流 %r' % name)
        return ole.openstream(name).read()
    finally:
        ole.close()


def mtef_from_ole(data):
    """OLE 复合文档 → MTEF 记录体（剥掉 OLE 壳 + 28B 头 + 版本头）"""
    raw = _ole_stream(data, 'Equation Native')
    return mtef_from_native(raw)


def mtef_from_native(raw):
    """Equation Native 流 → MTEF 记录体"""
    if len(raw) < 28:
        raise MTEFError('Equation Native 流太短')
    body = raw[28:]
    ver = body[0] if body else 0
    if ver == 5:
        # MathType：5 字节版本信息 + app key(cstring) + 1 字节 options
        p = 5
        end = raw.find(b'\x00', 28 + p)
        if end < 0:
            raise MTEFError('版本头 app key 无终止符')
        return raw[end + 2:]
    if ver in (2, 3):
        # v2/v3 = 微软公式编辑器 3.0（Equation Editor 3.0）的老格式。
        # 记录格式跟 MathType v5 不是一回事（没有 app key、记录字段也不同），
        # 按 v5 的规则去读必然错位。实测整个讲义库 8409 个公式里只有 22 个是 v3
        # （集中在 7 份文档），v5 那 8387 个是 100% 成功 —— 所以不做逆向，
        # 给一条能照做的提示即可。
        raise MTEFError(
            'MTEF v%d = 微软公式编辑器 3.0 老格式，本解析器只支持 MathType v5。'
            '用 Word 打开该文档，双击这个公式后选「公式 → 转换为新格式」'
            '（或全选后「公式 → 全部转换」），存盘后即可解析' % ver)
    raise MTEFError('不支持的 MTEF 版本 %d' % ver)


def equations_from_docx(docx_path):
    """从 docx 里取出所有 MathType 公式的 LaTeX

    返回 [{'part': 嵌入对象名, 'latex': ..., 'error': ...}, ...]
    """
    import zipfile
    out = []
    with zipfile.ZipFile(docx_path) as z:
        names = sorted(
            # 注意：word/embeddings/ 目录本身也是 zip 条目（0 字节、以 / 结尾），
            # 不过滤掉会被当成"内容为空的嵌入对象"，白报一堆假失败
            # （实测一份讲义库里就混进来 128 个这种条目）。
            (n for n in z.namelist()
             if n.startswith('word/embeddings/') and not n.endswith('/')),
            key=lambda n: int(''.join(c for c in n if c.isdigit()) or 0)
        )
        for n in names:
            rec = {'part': n.split('/')[-1], 'latex': '', 'tex_input': None,
                   'error': ''}
            try:
                body = mtef_from_ole(z.read(n))
                p = MTEFParser(body)
                node = p.parse()
                rec['latex'] = to_latex(node)
                # tex_input 是公式作者当初敲进去的原始 TeX（MathType 6+ 用
                # FUTURE 0x66 存）。它没有经过结构翻译的启发式猜测，通常更准，
                # 例如结构翻译会把 {{A}, {B}} 压平成 A,B，而 tex_input 保留了原样。
                # 调用方可以自己决定用哪个：能转 OMML 就优先用 tex_input。
                rec['tex_input'] = p.tex_input
                if not rec['latex']:
                    rec['error'] = '解析结果为空'
            except Exception as e:
                rec['error'] = '%s: %s' % (type(e).__name__, e)
            out.append(rec)
    return out


# ------------------------------------------------------------------ 自检

# 官方 SDK 文档里的示例流：-b±√(b²-4ac) / 2a  （MathType 4.0 生成）
# 用来验证 EQN_PREFS 边界、CHAR 编码、TMPL 结构、LINE_NULL 处理
SELFTEST_STREAM = bytes([
    0x05, 0x01, 0x00, 0x04, 0x00,
    0x44, 0x53, 0x4D, 0x54, 0x34, 0x00, 0x00,
    0x13, 0x57, 0x69, 0x6E, 0x41, 0x6C, 0x6C, 0x42, 0x61, 0x73, 0x69, 0x63,
    0x43, 0x6F, 0x64, 0x65, 0x50, 0x61, 0x67, 0x65, 0x73, 0x00,
    0x11, 0x05, 0x54, 0x69, 0x6D, 0x65, 0x73, 0x20, 0x4E, 0x65, 0x77,
    0x20, 0x52, 0x6F, 0x6D, 0x61, 0x6E, 0x00,
    0x11, 0x03, 0x53, 0x79, 0x6D, 0x62, 0x6F, 0x6C, 0x00,
    0x11, 0x05, 0x43, 0x6F, 0x75, 0x72, 0x69, 0x65, 0x72, 0x20, 0x4E, 0x65, 0x77, 0x00,
    0x11, 0x04, 0x4D, 0x54, 0x20, 0x45, 0x78, 0x74, 0x72, 0x61, 0x00,
    0x12, 0x00,
    0x08, 0x21, 0x2F, 0x45, 0x8F, 0x44, 0x2F, 0x41, 0x50, 0xF4, 0x10, 0x0F,
    0x47, 0x5F, 0x41, 0x50, 0xF2, 0x1F, 0x1E,
    0x41, 0x50, 0xF4, 0x15, 0x0F, 0x41, 0x00, 0xF4, 0x45, 0xF4, 0x25, 0xF4,
    0x8F, 0x42, 0x5F, 0x41, 0x00, 0xF4, 0x10, 0x0F, 0x43, 0x5F, 0x41, 0x00,
    0xF4, 0x8F, 0x45, 0xF4, 0x2A, 0x5F, 0x48, 0xF4, 0x8F, 0x41, 0x00, 0xF4,
    0x10, 0x0F, 0x40, 0xF4, 0x8F, 0x41, 0x7F, 0x48, 0xF4, 0x10, 0x0F, 0x41,
    0x2A, 0x5F, 0x44, 0x5F, 0x45, 0xF4, 0x5F, 0x45, 0xF4, 0x5F, 0x41, 0x0F,
    0x0C, 0x01, 0x00, 0x01, 0x00, 0x01, 0x02, 0x02, 0x02, 0x02, 0x00, 0x02,
    0x00, 0x01, 0x01, 0x01, 0x00, 0x03, 0x00, 0x01, 0x00, 0x04, 0x00, 0x00,
    0x0A,
    0x01, 0x00,
    0x03, 0x00, 0x0B, 0x00, 0x00,
    0x01, 0x00,
    0x02, 0x04, 0x86, 0x12, 0x22, 0x2D,
    0x02, 0x00, 0x83, 0x62, 0x00,
    0x02, 0x04, 0x86, 0xB1, 0x00, 0xB1,
    0x03, 0x00, 0x0A, 0x00, 0x00,
    0x01, 0x00,
    0x02, 0x00, 0x83, 0x62, 0x00,
    0x03, 0x00, 0x1C, 0x00, 0x00,
    0x0B, 0x01, 0x01,
    0x01, 0x00, 0x02, 0x00, 0x88, 0x32, 0x00, 0x00,
    0x00,
    0x0A, 0x02, 0x04, 0x86, 0x12, 0x22, 0x2D,
    0x02, 0x00, 0x88, 0x34, 0x00,
    0x02, 0x00, 0x83, 0x61, 0x00,
    0x02, 0x00, 0x83, 0x63, 0x00,
    0x00,
    0x0B, 0x01, 0x01,
    0x00,
    0x00,
    0x0A, 0x01, 0x00, 0x02, 0x00, 0x88, 0x32, 0x00, 0x02, 0x00, 0x83, 0x61, 0x00, 0x00,
    0x00,
    0x00,
])

SELFTEST_EXPECT = '\\frac{-b\\pm \\sqrt{b^{2}-4ac}}{2a}'


def _norm(s):
    """归一化：去掉空格和花括号（b^2 与 b^{2} 等价）"""
    for ch in (' ', '{', '}'):
        s = s.replace(ch, '')
    return s


def selftest():
    print('=== MTEF 自检（官方 SDK 示例流：-b±√(b²-4ac)/2a）===\n')
    ok = True

    # 前 12 字节是 MTEF 版本头（5 字节版本信息 + "DSMT4\0" + 1 字节 options），
    # 真实数据里由 mtef_from_native() 剥掉，这里手动剥
    node = MTEFParser(SELFTEST_STREAM[12:]).parse()
    got = to_latex(node)
    print('  解析结果: %s' % got)
    print('  期望    : %s' % SELFTEST_EXPECT)
    if _norm(got) == _norm(SELFTEST_EXPECT):
        print('  PASS  官方示例流')
    else:
        print('  FAIL  官方示例流')
        ok = False

    return ok


# ------------------------------------------------------------------ CLI

def main():
    ap = argparse.ArgumentParser(description='MTEF v5 → LaTeX')
    ap.add_argument('target', nargs='?', help='.docx 文件')
    ap.add_argument('--json', help='导出 JSON')
    ap.add_argument('--selftest', action='store_true', help='跑自检')
    args = ap.parse_args()

    if args.selftest:
        sys.exit(0 if selftest() else 1)

    if not args.target:
        ap.print_help()
        sys.exit(1)

    eqs = equations_from_docx(args.target)
    for r in eqs:
        if r['error']:
            print('%-22s [错误] %s' % (r['part'], r['error']))
        else:
            print('%-22s %s' % (r['part'], r['latex']))

    good = sum(1 for r in eqs if not r['error'] and r['latex'])
    print('\n共 %d 个公式，成功 %d，失败 %d' % (len(eqs), good, len(eqs) - good))

    if args.json:
        with open(args.json, 'w', encoding='utf-8') as f:
            json.dump(eqs, f, ensure_ascii=False, indent=1)
        print('已导出 %s' % args.json)


if __name__ == '__main__':
    main()
