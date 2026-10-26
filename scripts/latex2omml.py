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


def _wrap_bare_bars(mml_root):
    r"""
    把同一个 mrow 里「成对的裸竖线 |」显式包进 mrow（原地修改）。

    病根：`|a||b|` 的 MathML 是 mo(|) mi(a) mo(|) mo(|) mi(b) mo(|)，
    相邻两个 mo(|) 让 MML2OMML.XSL 误判成一个"三分组定界符"
    （begChr='|' endChr='|' + 两条 sep），把 a、b 全吞掉，
    最终产出 `<m:d><m:e/><m:e/><m:e/></m:d>` —— 三个空壳，Word 里一片空白，
    而且**不报错**，所以降级机制完全抓不到。

    修法：按出现顺序把裸竖线两两配对，配成
        <mrow><mo form=prefix>|</mo>…<mo form=postfix>|</mo></mrow>
    与 `\left|a\right|` 的 MathML 结构一致，XSL 便能正确生成定界符组。

    保守策略：只处理**全是裸竖线**（无 form 属性）的 mrow，且个数为偶数；
    已经写过 \left| / \right| 的、奇数个（可能嵌套）的一律不动，保持原行为。
    """
    MML = '{http://www.w3.org/1998/Math/MathML}'

    for mrow in list(mml_root.iter(MML + 'mrow')):
        kids = list(mrow)
        if not kids:
            continue
        # 找出「裸竖线」：文本是 | 且没有被 \left/\right 标记过
        bare_idx = []
        for i, k in enumerate(kids):
            if k.tag != MML + 'mo':
                continue
            if (k.text or '').strip() != '|':
                continue
            if k.get('form') is not None or k.get('fence') == 'true':
                continue          # \left|...\right| 的产物，本来就对
            bare_idx.append(i)

        if len(bare_idx) < 4 or len(bare_idx) % 2:
            continue              # 单纯 |x| （2 条）XSL 本来就能处理，不折腾
        # 必须真的存在「相邻两条」（否则是 |a|+|b| 这种，XSL 也没问题），
        # 但为稳妥：只要这一层有裸竖线成对，就统一显式化，行为等价、更保险
        pairs = list(zip(bare_idx[0::2], bare_idx[1::2]))

        out = []
        pair_at = {i: (i, j) for i, j in pairs}
        close_at = {j: i for i, j in pairs}
        i = 0
        while i < len(kids):
            if i in pair_at:
                _, j = pair_at[i]
                wrapper = etree.Element(MML + 'mrow')
                op, cl = kids[i], kids[j]
                op.set('stretchy', 'true')
                op.set('fence', 'true')
                op.set('form', 'prefix')
                cl.set('stretchy', 'true')
                cl.set('fence', 'true')
                cl.set('form', 'postfix')
                wrapper.append(op)
                for k in range(i + 1, j):
                    wrapper.append(kids[k])
                wrapper.append(cl)
                out.append(wrapper)
                i = j + 1
            elif i in close_at:
                # 落单的闭合符号（理论上不会走到这里），原样保留
                out.append(kids[i])
                i += 1
            else:
                out.append(kids[i])
                i += 1

        for c in list(mrow):
            mrow.remove(c)
        for c in out:
            mrow.append(c)

    return mml_root


def preprocess_mathml(mml_root):
    """
    XSL 转换前的 MathML 修补（原地修改，返回同一个 root）：
    1. 符号码位修正（\\perp 等 latex2mathml 的映射错误）
    2. 给 accent 类 mover/munder 补 accent 属性
    3. 裸竖线两两配对，避免 |a||b| 被 XSL 吞成空公式
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

    # --- 3. 裸竖线配对（|a||b| 这类，见 _wrap_bare_bars 的说明）
    _wrap_bare_bars(mml_root)

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
    # 注意替换串必须是**单反斜杠**：premap 用 lambda 原样输出，
    # 写成 r'\\mathrm{\1}' 会得到字面的 "\\mathrm{\1}"（历史 bug，2026-10 修）
    (r'\\unit\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}', r'\mathrm{\1}'),
    (r'\\si\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}', r'\mathrm{\1}'),
    # 函数名：latex2mathml 认识 \sin \cos \log \max，但不认识 \arg，
    # 会**静默**把 "\arg" 当普通文字吐进公式（Word 里显示成斜体 "\arg"）。
    # 负向断言避免误伤 \argmax 这类自定义命令。
    # ⚠️ 别用 \operatorname{arg}：它生成 <mo>arg</mo>，XSL 会把它当成**定界符**，
    #    产出 begChr="arg" 的畸形 m:d（实测 2026-10）；\mathrm 生成 <mi>，安全。
    (r'\\arg(?![A-Za-z])', r'\mathrm{arg}'),
    # \vspace 在公式里没有意义，且会把 "\vspace1em" 文本漏进公式，直接剔除
    (r'\\vspace\*?\{[^{}]*\}', ''),
]
_PREMAP_COMPILED = [(re.compile(p), r) for p, r in _LATEX_PREMAP]

# 替换串里的 \1 \2 组引用
_GROUP_REF = re.compile(r'\\(\d)')


def _expand_repl(repl: str, m) -> str:
    """把替换串里的 \\1 换成对应捕获组，其余反斜杠原样保留"""
    if '\\' not in repl:
        return repl
    return _GROUP_REF.sub(lambda g: m.group(int(g.group(1))) or '', repl)


def premap_latex(latex: str) -> str:
    """转换前把扩展/单位命令重写为核心 LaTeX 命令"""
    for pat, repl in _PREMAP_COMPILED:
        # 用 lambda 传替换串：replacement 里的 \circ 等按字面输出，不经转义解析；
        # \1 这类组引用在 _expand_repl 里手动展开
        latex = pat.sub(lambda m, r=repl: _expand_repl(r, m), latex)
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


# ---------------------------------------------------------------- 错误人话化
class MathConvertError(ValueError):
    """单个公式转换失败。除了技术信息，还带「原因 + 怎么办」的人话说明。"""

    def __init__(self, latex, cause='', reason='', hint=''):
        self.latex = latex
        self.cause = cause
        self.reason = reason
        self.hint = hint
        super().__init__(reason or cause or '公式转换失败')

    def human(self, where=''):
        loc = f'位置 {where} · ' if where else ''
        out = [f'[公式没转成] {loc}原文: {self.latex}']
        if self.reason:
            out.append(f'  原因：{self.reason}')
        if self.cause:
            out.append(f'  技术细节：{self.cause}')
        out.append('  影响：这一处已按原样保留为文字，文档其余部分不受影响')
        if self.hint:
            out.append(f'  怎么办：{self.hint}')
        return '\n'.join(out)


# 引擎支持范围之外的命令：latex2mathml 不报错，直接把命令名当文本吐出来
_LATEX_LEFTOVER = re.compile(r'\\([A-Za-z]{2,})')


def _diagnose(latex, omml_text, exc=None):
    """根据失败现象归纳「原因 + 怎么办」。返回 (reason, hint, cause)"""
    if exc is not None:
        msg = f'{type(exc).__name__}: {exc}'
        low = str(exc).lower()
        if 'xsl' in low or '文件' in str(exc) and '找不到' in str(exc):
            return ('找不到 MML2OMML.XSL（公式转换表），引擎没法工作',
                    '确认 scripts/ 目录下有 MML2OMML.XSL；没有就重装本 skill，'
                    '或运行 python scripts/doctor.py 看缺什么', msg)
        if 'xmlsyntax' in low or 'syntax' in low:
            return ('这段 LaTeX 本身不合法（括号/环境没闭合最常见）',
                    '数一下花括号 {} 和 \\left / \\right 是否成对、'
                    '\\begin{...} 有没有对应的 \\end{...}', msg)
        return ('转换引擎内部报错', '把这一行的报错贴给维护者；临时可改用更简单的写法', msg)

    if not omml_text.strip():
        return ('转换结果是一片空白（公式内容被吞了）',
                '这是已知的边界情况，见 SKILL.md「常见坑」；'
                '可临时把 |a||b| 写成 \\left|a\\right|\\left|b\\right|', '空公式')

    # 残留 LaTeX 命令 = 该命令不被支持，被当普通文字写进了公式。
    # 注意 OMML 里命令名可能和后面的变量粘在一起（\foo{x} -> "\foox"），
    # 所以拿源 LaTeX 里的命令集去对：谁是它的前缀，谁就是没认出来的那个。
    leftover = _LATEX_LEFTOVER.search(omml_text)
    if leftover:
        raw = leftover.group(1)
        src_cmds = set(re.findall(r'\\([A-Za-z]+)', latex))
        name = '\\' + raw
        for c in sorted(src_cmds, key=len, reverse=True):
            if raw.startswith(c):
                name = '\\' + c
                break
        common = {
            '\\text': '试试 \\mathrm{...} 或 \\mbox{...}',
            '\\mbox': '试试 \\text{...} 或 \\mathrm{...}',
            '\\foo': '把命令换成引擎支持的等价写法',
        }.get(name, '')
        return (f'LaTeX 命令 {name} 不被支持，被当成普通文字写进了公式',
                common or '在 SKILL.md 的「命令支持表」里找等价写法，'
                          '或用 \\mathrm{} / \\text{} 包住这段内容',
                f'残留命令 {name}')

    return ('未知原因', '见 SKILL.md「常见坑」', '')


def convert_or_explain(latex, display='inline'):
    """转换公式；失败或出现静默损坏时，抛带人话的 MathConvertError。"""
    try:
        omml = latex_to_omml(latex, display)
    except Exception as e:                                  # noqa: BLE001
        reason, hint, cause = _diagnose(latex, '', e)
        raise MathConvertError(latex, cause, reason, hint) from e

    root = etree.fromstring(omml)
    body = ''.join(t.text or '' for t in root.iter(M + 't'))
    reason, hint, cause = _diagnose(latex, body, None)
    # 原文本就没字符（如 \, ）时不追究空公式
    if not body.strip() and re.search(r'[A-Za-z0-9\u4e00-\u9fff]', latex):
        raise MathConvertError(latex, cause, reason, hint)
    if cause.startswith('残留命令'):
        raise MathConvertError(latex, cause, reason, hint)
    return omml





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


def add_paragraph(doc, text, style=None, on_error='text', where='', errors=None):
    """
    往 docx 段落里写含 $...$ 公式的文本，公式转成原生 OMML 对象。

    on_error: 'text'  公式转换失败时降级为纯文本（推荐，不中断）
              'raise' 直接抛异常
    where:    出错时用来定位的说明，如 '第 12 段' / '第 3 页第 2 行'
    errors:   传一个 list，失败信息会追加进去，方便上层汇总
    """
    from docx.oxml import parse_xml

    p = doc.add_paragraph(style=style)
    for seg in split_math(text):
        if seg['type'] == 'text':
            if seg['value']:
                p.add_run(seg['value'])
            continue
        try:
            omml = convert_or_explain(
                seg['value'],
                'block' if seg['type'] == 'block' else 'inline'
            )
            p._p.append(parse_xml(omml))
        except MathConvertError as e:
            if on_error == 'raise':
                raise
            p.add_run(f"[{seg['value']}]")
            sys.stderr.write(e.human(where) + '\n')
            if errors is not None:
                errors.append(e)
        except Exception as e:                              # noqa: BLE001
            if on_error == 'raise':
                raise
            # 兜底：连诊断都失败的极端情况，保证文档不丢内容
            p.add_run(f"[{seg['value']}]")
            loc = f'位置 {where} · ' if where else ''
            sys.stderr.write(
                f'[公式没转成] {loc}原文: {seg["value"]}\n'
                f'  原因：{type(e).__name__}: {e}\n'
                f'  影响：这一处已按原样保留为文字，文档其余部分不受影响\n'
                f'  怎么办：把这一行报错贴给维护者\n'
            )
            if errors is not None:
                errors.append(e)
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
        # 回归：紧邻绝对值 |a||b| 曾被 XSL 吞成三个空 <m:e/>，Word 里一片空白
        (r'|a_{i}||a_{j}|', 'm:sSub', '紧邻绝对值'),
        # 回归：分段函数右边不能凭空多一个 ")"
        (r'\begin{cases}1 & x>0\\ 0 & x\leq 0\end{cases}', '<m:endChr m:val=""/>',
         '分段函数右定界符'),
    ]
    ok = fail = 0
    detail = []
    for latex, needle, name in cases:
        try:
            xml = latex_to_omml(latex).decode('utf-8')
            hit = needle in xml
            # 附加约束：公式里必须有真内容（防「结构对了但内容被吞」）
            root = etree.fromstring(xml.encode('utf-8'))
            body = ''.join(t.text or '' for t in root.iter(M + 't'))
            if hit and name in ('紧邻绝对值',):
                hit = ('a' in body) and ('i' in body) and ('j' in body)
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
