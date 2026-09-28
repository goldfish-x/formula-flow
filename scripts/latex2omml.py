# -*- coding: utf-8 -*-
"""
latex2omml — LaTeX 转 Word 原生可编辑公式（OMML）核心引擎

链路：  LaTeX --latex2mathml--> MathML --MML2OMML.XSL--> OMML --> docx

生成的公式是 Word 原生公式对象（oMath），可在 Word 里双击编辑、
改字号、改颜色，而不是图片。

用法（作为库）：
    from latex2omml import latex_to_omml, build_docx
    omml_bytes = latex_to_omml(r"\\frac{1}{6}")
    build_docx([("标题", "Heading 1"), ("正文 $x^2$", None)], "out.docx")
"""
import os
import re
import sys

from lxml import etree
import latex2mathml.converter as _conv

HERE = os.path.dirname(os.path.abspath(__file__))

# XSL 查找顺序：skill 内置 > Office 安装目录（32/64 位）
XSL_CANDIDATES = [
    os.path.join(HERE, 'MML2OMML.XSL'),
    r'C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL',
    r'C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL',
    r'C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL',
]

MML_NS = 'http://www.w3.org/1998/Math/MathML'
OMML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
M = '{' + OMML_NS + '}'

_transform = None


# ---------------------------------------------------------------- OMML 后处理

import re as _re

# 匹配 oMath 第一个子节点是 <m:r><m:t>f(x)=...</m:t></m:r> 这种「函数名+括号+等号」模式
_LEADING_FUNCEQ = _re.compile(
    r'<m:r>\s*<m:t>\s*(?P<name>[A-Za-zΑ-Ωα-ω]+)\s*\(\s*(?P<arg>[^)]+?)\s*\)\s*=\s*</m:t>\s*</m:r>'
)


# ---------------------------------------------------------------- MathML 预处理

# 「上方/下方标注」字符：这些出现在 mover/munder 里时是 accent（紧贴主体的
# 强调符号），而不是 lim（悬空的上/下限）。缺了 accent 属性，XSL 会把它们
# 转成 m:limUpp/m:limLow，视觉上符号悬空、间距过大，完全不是日常书写习惯。
_ACCENT_CHARS = set(
    '\u2192'   # →  右箭头      \vec \overrightarrow
    '\u2190'   # ←  左箭头      \overleftarrow
    '\u2194'   # ↔  双向箭头    \overleftrightarrow
    '\u20d7'   # ⃗  组合右箭头
    '\u20d6'   # ⃖  组合左箭头
    '\u20e1'   # ⃡  组合左右箭头
    '\u0302'   # ̂  组合抑扬符  \hat \widehat
    '\u02c6'   # ˆ  抑扬符
    '\u0303'   # ̃  组合波浪号  \tilde
    '\u02dc'   # ˜  波浪号      \widetilde
    '\u0307'   # ̇  组合上点    \dot
    '\u02d9'   # ˙  上点
    '\u0308'   # ̈  组合分音符  \ddot
    '\u00a8'   # ¨  分音符
    '\u030c'   # ̌  组合抑扬符  \check
    '\u02c7'   # ˇ  抑扬符(倒)
    '\u0301'   # ́  组合锐音符  \acute
    '\u00b4'   # ´  锐音符
    '\u0300'   # ̀  组合钝音符  \grave
    '\u0060'   # `  钝音符
    '\u0306'   # ̆  组合短音符  \breve
    '\u02d8'   # ˘  短音符
    '\u030a'   # ̊  组合上圆圈  \mathring
    '\u02da'   # ˚  上圆圈
    '\u00af'   # ¯  长音符      \bar \overline
    '\u0304'   # ̄  组合长音符
    '\u0332'   # ̲  组合下横线  \underline
    '\u0331'   # ̱  组合下长音符
    # --- ASCII 变体：latex2mathml 对部分命令直接吐 ASCII 码位
    '\u007e'   # ~  波浪号      \tilde \widetilde
    '\u005e'   # ^  抑扬符      \widehat
    '\u005f'   # _  下横线      \underline
    '\u2015'   # ―  水平横线    \underline
    '\u2013'   # –  短横线
    '\u2014'   # —  长横线
)

# latex2mathml 符号表错误：\perp 应输出 U+22A5(⊥ UP TACK)，
# 实际输出 U+27C2(⟂)，后者是「竖线加横stroke」，不是垂直符号。
_MATHML_CHAR_FIX = {
    '\u27c2': '\u22a5',   # ⟂ -> ⊥  \perp
}


def preprocess_mathml(mml_root):
    """
    XSL 转换前的 MathML 修补（原地修改，返回同一个 root）：
    1. 符号码位修正（\\perp 等 latex2mathml 的映射错误）
    2. 给 accent 类 mover/munder 补 accent 属性
    """
    MML = '{http://www.w3.org/1998/Math/MathML}'

    # --- 1. 码位修正：遍历所有含文本的元素
    for el in mml_root.iter():
        if el.text:
            new = ''.join(_MATHML_CHAR_FIX.get(ch, ch) for ch in el.text)
            if new != el.text:
                el.text = new

    # --- 2. accent 属性补全
    def _is_accent(node):
        """判断节点是否是 accent 字符（纯文本且字符在 accent 集合里）"""
        if node is None:
            return False
        # 取节点全部文本（可能嵌在 mi/mo/mrow 里）
        txt = ''.join(node.itertext())
        txt = txt.strip()
        return bool(txt) and all(ch in _ACCENT_CHARS for ch in txt)

    for tag, attr in (('mover', 'accent'), ('munder', 'accentunder')):
        for node in mml_root.iter(MML + tag):
            kids = [k for k in node if k.tag != MML + 'mstyle']
            if len(kids) < 2:
                continue
            # mover(base, over) — 第二个子元素是「上方字符」
            # munder(base, under) — 同理
            if _is_accent(kids[1]):
                node.set(attr, 'true')

    return mml_root


def postprocess_omml(omml_xml: str) -> str:
    r"""
    OMML 字符串层修复：

    XSL 会把「函数名(自变量)=」合并成一个 <m:r> 文本，Word 自动给那个
    孤立的「(」配一个撑大的「)」。把它拆成 f + (x) + = 三个数学对象。

    ⚠️ 这里曾经还有一段「删除 cases 的 <m:endChr m:val=""/>」的修复，是**错的**：
    MML2OMML.XSL 的规则是「定界符等于默认值时才省略」，OMML 里
    m:begChr 默认 '('、m:endChr 默认 ')'。所以：

      \\left( ... \\right)        -> XSL 不写 begChr/endChr   （默认即 '(' ')' ，正确）
      A=[-2,4)                   -> XSL 只写 begChr='['      （endChr 省略＝默认 ')' ，正确）
      \\begin{cases}             -> XSL 写 begChr='{' + endChr=''（空串＝不画右括号，正确）

    把 cases 那个显式的 <m:endChr m:val=""/> 删掉，Word 就会退回默认的 ')'，
    于是分段函数右边凭空多出一个撑大的右圆括号。**绝不能删**。
    """
    s = omml_xml

    # 修复：拆掉开篇的 f(x)= 文本 run
    def _split_leading(m):
        name, arg = m.group('name'), m.group('arg')
        # 用 OML 命名空间前缀 m:，并保留 xmlns 已在根上声明
        return (
            f'<m:r><m:t>{name}</m:t></m:r>'
            f'<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr>'
            f'<m:e><m:r><m:t>{arg}</m:t></m:r></m:e></m:d>'
            f'<m:r><m:t>=</m:t></m:r>'
        )

    s = _LEADING_FUNCEQ.sub(_split_leading, s, count=1)

    return s


# 下划线类字符：\underline 经 XSL 后是 m:limLow + 这些字符，
# 需要转成 Word 的 m:bar(pos=bot) 才是「紧贴主体的下划线」
_UNDERLINE_CHARS = set('\u2015\u0332\u005f\u2013\u2014\u00af')


def postprocess_omml_root(root):
    r"""
    OMML 结构化修复（lxml Element 层面，处理正则啃不动的嵌套情形）：
    把 \underline 退化出的 m:limLow + 横线 转回 Word 的 m:bar(pos=bot)。
    """
    # 从 root 里取实际命名空间前缀（XSL 输出可能带前缀也可能是默认前缀 m:）
    tag = root.tag
    ns = tag[tag.find('{') + 1:tag.find('}')] if '{' in tag else OMML_NS
    M = '{%s}' % ns

    def q(name):
        return M + name

    for lim in list(root.iter(q('limLow'))):
        e = lim.find(q('e'))
        limel = lim.find(q('lim'))
        if e is None or limel is None:
            continue
        txt = ''.join(limel.itertext()).strip()
        if not txt or not all(ch in _UNDERLINE_CHARS for ch in txt):
            continue
        # 构造 <m:bar><m:barPr><m:pos m:val="bot"/></m:barPr><m:e>..</m:e></m:bar>
        bar = etree.Element(q('bar'))
        barPr = etree.SubElement(bar, q('barPr'))
        pos = etree.SubElement(barPr, q('pos'))
        pos.set(q('val'), 'bot')
        bar.append(e)
        parent = lim.getparent()
        if parent is not None:
            parent.replace(lim, bar)
    return root


# ---------------------------------------------------------------- 正体化（数字/括号不斜体）
# GB 3102 / ISO 80000 排版惯例：数字、括号、运算符、标点用正体，变量字母用斜体。
# 背景：Word 对「无 m:sty 的裸 run」按字符类别自动排版（数字/括号自动正体），
# 但 WPS 的排版引擎不做自动判定，裸 run 一律按斜体渲染，数字和括号看着就是斜的。
# 因此：不含字母的文本段显式写 m:sty="p"；含字母的段保持默认（Word/WPS 都排斜体）。

_XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


def _char_is_letter(ch):
    """字母（拉丁/希腊/双算体等，Unicode 类别 L*）→ 保持默认斜体；其余 → 正体"""
    import unicodedata
    return unicodedata.category(ch).startswith('L')


def _split_letter_segments(text):
    """把文本切成 [(是否字母段, 片段), ...]，同类字符合并成一段"""
    segs = []
    for ch in text:
        cls = _char_is_letter(ch)
        if segs and segs[-1][0] == cls:
            segs[-1][1] += ch
        else:
            segs.append([cls, ch])
    return [(cls, s) for cls, s in segs]


def _make_run(M, text, upright, base_rPr=None):
    """构造一个 m:r；upright=True 时写 m:sty="p"，base_rPr 为要继承的原 rPr"""
    r = etree.Element(M + 'r')
    rPr = None
    if base_rPr is not None:
        rPr = etree.fromstring(etree.tostring(base_rPr))
        r.append(rPr)
    if upright:
        if rPr is None:
            rPr = etree.SubElement(r, M + 'rPr')
        sty = etree.SubElement(rPr, M + 'sty')
        sty.set(M + 'val', 'p')
    t = etree.SubElement(r, M + 't')
    t.text = text
    if text != text.strip():
        t.set(_XML_SPACE, 'preserve')
    return r


def upright_nonletters(root):
    """把 OMML 里数字/括号/运算符/标点所在的 run 标记为正体（m:sty="p"）。

    - 整个 run 无字母：直接给 run 加 sty="p"
    - run 内字母与非字母混排（如 f(x)=3x+5）：拆成多个 run，非字母段 sty="p"
    - 已有显式 m:sty 或 m:nor 的 run：尊重原作者意图，跳过不动
    """
    tag = root.tag
    ns = tag[tag.find('{') + 1:tag.find('}')] if '{' in tag else OMML_NS
    M = '{%s}' % ns

    for r in list(root.iter(M + 'r')):
        rPr = r.find(M + 'rPr')
        if rPr is not None and (rPr.find(M + 'sty') is not None
                                or rPr.find(M + 'nor') is not None):
            continue
        t = r.find(M + 't')
        if t is None:
            continue
        text = t.text or ''
        if not text:
            continue
        segs = _split_letter_segments(text)
        has_letter = any(cls for cls, _ in segs)
        if not has_letter:
            # 整段非字母：原地加 sty="p"
            if rPr is None:
                rPr = etree.Element(M + 'rPr')
                r.insert(0, rPr)
            sty = etree.SubElement(rPr, M + 'sty')
            sty.set(M + 'val', 'p')
            continue
        if len(segs) == 1:
            continue
        # 混排：拆成多个 run
        new_runs = [_make_run(M, seg, upright=not cls, base_rPr=rPr)
                    for cls, seg in segs]
        parent = r.getparent()
        if parent is None:
            continue
        idx = parent.index(r)
        for i, nr in enumerate(new_runs):
            parent.insert(idx + i, nr)
        parent.remove(r)
    return root


def _extract_bundled_xsl():
    """从内嵌的 xsl_bundle（gzip+base64）解出 MML2OMML.XSL。

    用于受限分发场景：上传平台不允许 .xsl 扩展名时，包里没有
    MML2OMML.XSL 文件，运行时从这里自动还原。优先写到 scripts/ 目录
    （可写则下次直接命中 XSL_CANDIDATES），不可写则退到系统临时目录。
    """
    try:
        import base64 as _b64
        import gzip as _gzip
        import tempfile as _tmp
        import xsl_bundle
    except ImportError:
        return None
    data = _gzip.decompress(_b64.b64decode(xsl_bundle.XSL_GZ_B64))
    try:
        target = os.path.join(HERE, 'MML2OMML.XSL')
        with open(target, 'wb') as f:
            f.write(data)
        return target
    except OSError:
        fd, tmp = _tmp.mkstemp(suffix='.xsl')
        with os.fdopen(fd, 'wb') as f:
            f.write(data)
        return tmp


def _get_transform():
    """懒加载并缓存 XSLT 转换器"""
    global _transform
    if _transform is None:
        p = None
        for cand in XSL_CANDIDATES:
            if os.path.exists(cand):
                p = cand
                break
        if p is None:
            p = _extract_bundled_xsl()
        if p is None:
            raise RuntimeError(
                '找不到 MML2OMML.XSL。请确认 Microsoft Office 已安装，'
                '或把 MML2OMML.XSL 放到 scripts/ 目录下。'
            )
        _transform = etree.XSLT(etree.parse(p))
    return _transform


# ---------------------------------------------------------------- LaTeX 预映射
# latex2mathml 对下列命令「静默失败」：不报错、直接把命令名当文本吐进公式
# （如 \celsius 原样出现在结果里）。这里在进入转换前重写成等价的核心命令。
_LATEX_PREMAP = [
    (r'\\degreeCelsius\b', r'^{\circ}\mathrm{C}'),
    (r'\\celsius\b', r'^{\circ}\mathrm{C}'),
    (r'\\degree\b', r'^{\circ}'),
    (r'\\arcmin\b', r"^{\prime}"),
    (r'\\arcsec\b', r"^{\prime\prime}"),
    (r'\\ohm\b', r'\Omega'),
    (r'\\micro\b', r'\mu'),
    (r'\\permil\b', r'\text{‰}'),
    (r'\\AA\b', r'\text{Å}'),
    # 微积分：varoiint 在 latex2mathml 会静默泄漏，改用原生 \oiint (U+222F)
    # 注意不能用 \b：`\varoiint_` 中 `_` 也是词字符，\b 会失配，用负向断言
    (r'\\varoiint(?![A-Za-z])', r'\oiint'),
    (r'\\varoiiint(?![A-Za-z])', r'\oiiint'),
    # \unit{m/s^2} -> \mathrm{m/s^2}（允许一层嵌套花括号）
    (r'\\unit\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}', r'\\mathrm{\1}'),
    (r'\\si\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}', r'\\mathrm{\1}'),
]
_PREMAP_COMPILED = [(re.compile(p), r) for p, r in _LATEX_PREMAP]


def premap_latex(latex: str) -> str:
    """转换前把扩展/单位命令重写为核心 LaTeX 命令"""
    for pat, repl in _PREMAP_COMPILED:
        # 用 lambda 传替换串：replacement 里的 \\circ 等按字面输出，不经转义解析
        latex = pat.sub(lambda m, r=repl: r, latex)
    return latex


def latex_to_mathml(latex: str) -> str:
    """LaTeX -> MathML 字符串"""
    return _conv.convert(premap_latex(latex))


def latex_to_omml(latex: str, display: str = 'inline') -> bytes:
    """
    LaTeX -> OMML XML（bytes），可直接插入 docx 段落。

    display: 'inline' 行内公式 | 'block' 独立成行公式
    """
    mml = _conv.convert(premap_latex(latex))
    try:
        doc = etree.fromstring(mml.encode('utf-8'))
    except etree.XMLSyntaxError as e:
        raise ValueError(f'LaTeX 转 MathML 失败: {latex!r} ({e})') from e
    doc.set('display', display)
    # MathML 预处理：修正符号码位、给 accent 类 mover/munder 补 accent 属性
    doc = preprocess_mathml(doc)
    result = _get_transform()(doc)
    root = result.getroot()
    etree.cleanup_namespaces(root)
    # 先做字符串级修复（f(x)= 拆分会产生新的裸 run，必须在其后做正体化）
    xml = etree.tostring(root, encoding='unicode')
    xml = postprocess_omml(xml)
    # 重新解析，再做结构级修复：
    #   1) \underline 退化的 m:limLow 转回 m:bar(pos=bot)
    #   2) 数字/括号/运算符/标点正体化（m:sty="p"，Word/WPS 渲染一致）
    root = etree.fromstring(xml.encode('utf-8'))
    root = postprocess_omml_root(root)
    root = upright_nonletters(root)
    # 注意：tostring 默认 ASCII 编码会把 ⊆ ∅ 等符号转成 &#8838; 实体，
    # 必须用 unicode 编码保留真实字符，否则 Word 打开后符号可能异常。
    xml = etree.tostring(root, encoding='unicode')
    return xml.encode('utf-8')


# ---------------------------------------------------------------- 段落构建

# 匹配 $$...$$（独立）优先于 $...$（行内）
_MATH_PATTERN = re.compile(r'(\$\$.+?\$\$|\$.+?\$)', re.S)


def split_math(text: str):
    """
    把文本切成片段列表，每项为 dict:
      {'type': 'text', 'value': str}
      {'type': 'inline', 'value': latex}
      {'type': 'block',  'value': latex}
    """
    out = []
    for part in _MATH_PATTERN.split(text):
        if not part:
            continue
        if part.startswith('$$') and part.endswith('$$') and len(part) > 4:
            out.append({'type': 'block', 'value': part[2:-2].strip()})
        elif part.startswith('$') and part.endswith('$') and len(part) > 2:
            out.append({'type': 'inline', 'value': part[1:-1].strip()})
        else:
            out.append({'type': 'text', 'value': part})
    return out


def add_paragraph(doc, text, style=None, on_error='text'):
    """
    往 docx 段落里写含 $...$ 公式的文本，公式转成原生 OMML 对象。

    on_error: 'text'  公式转换失败时降级为纯文本（推荐，不中断）
              'raise' 直接抛异常
    """
    from docx.oxml import parse_xml

    p = doc.add_paragraph(style=style)
    for seg in split_math(text):
        if seg['type'] == 'text':
            if seg['value']:
                p.add_run(seg['value'])
            continue
        try:
            omml = latex_to_omml(
                seg['value'],
                'block' if seg['type'] == 'block' else 'inline'
            )
            p._p.append(parse_xml(omml))
        except Exception as e:
            if on_error == 'raise':
                raise
            # 降级：保留原始 LaTeX 文本，保证文档不丢内容
            p.add_run(f"[{seg['value']}]")
            sys.stderr.write(f'[latex2omml] 公式转换失败，已降级为文本: '
                             f'{seg["value"]!r} -> {e}\n')
    return p


def build_docx(blocks, out_path, title=None):
    """
    blocks: [(text, style), ...] style 可为 None 或 'Heading 1' 等
            text 中可用 $...$ 写行内公式、$$...$$ 写独立公式
    """
    from docx import Document

    doc = Document()
    if title:
        doc.add_heading(title, level=0)
    for text, style in blocks:
        add_paragraph(doc, text, style=style)
    doc.save(out_path)
    return out_path


# ---------------------------------------------------------------- 自检

def selftest():
    """跑一遍核心符号，返回 (通过数, 失败数, 明细)"""
    cases = [
        (r'\frac{1}{6}', 'm:f', '分数'),
        (r'x^{2}', 'm:sSup', '上标'),
        (r'a_{i}', 'm:sSub', '下标'),
        (r'A\subseteq B', '⊆', '子集'),
        (r'A\subsetneqq B', '⫋', '真子集'),
        (r'\varnothing', '∅', '空集'),
        (r'x\in\mathbb{N}', '∈', '属于+黑板粗体'),
        (r'\sqrt{x}', 'm:rad', '根号'),
        (r'\sum_{i=1}^{n}', 'm:nary', '求和'),
        (r'\{x\mid x>0\}', '{', '集合构造'),
        (r'\alpha+\beta', 'α', '希腊字母'),
        (r'\int_{0}^{1}f(x)dx', 'm:nary', '积分'),
    ]
    ok = fail = 0
    detail = []
    for latex, needle, name in cases:
        try:
            xml = latex_to_omml(latex).decode('utf-8')
            hit = needle in xml
        except Exception as e:
            hit = False
            xml = str(e)
        detail.append((name, latex, hit))
        if hit:
            ok += 1
        else:
            fail += 1
    return ok, fail, detail


if __name__ == '__main__':
    ok, fail, detail = selftest()
    for name, latex, hit in detail:
        print(f'{"PASS" if hit else "FAIL"}  {name:14s} {latex}')
    print(f'\n通过 {ok} / {ok + fail}')
    sys.exit(0 if fail == 0 else 1)
