# -*- coding: utf-8 -*-
"""
upgrade_math — 把 docx 里「Unicode 硬拼的伪公式」升级为 Word 原生公式

背景：很多讲义的公式是用 Unicode 字符硬拼出来的（a²+b² ≥ 2ab、A⊆B、
√(ab)、(a+b)/2），看着像公式，实际是纯文本——不能编辑、排版不专业、
上标对齐差。本工具把它们转成真正的 oMath 对象。

设计原则：**默认只读，不修改原文件**。先出报告，人工确认后再 --apply。

用法：
    # 1. 出报告（不动原文件）
    python upgrade_math.py "第4讲.docx"

    # 2. 看转换结果抽样
    python upgrade_math.py "第4讲.docx" --sample 15

    # 3. 确认后应用（自动备份 .bak.docx，输出 _公式升级.docx）
    python upgrade_math.py "第4讲.docx" --apply

    # 4. 只转高置信度（推荐首次使用）
    python upgrade_math.py "第4讲.docx" --apply --min-level high
"""
import argparse
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# ---------------------------------------------------------------- 字符表

UNI_SUP = {
    '⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
    '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9',
    'ⁿ': 'n', '⁺': '+', '⁻': '-', 'ⁱ': 'i', 'ʲ': 'j',
}
UNI_SUB = {
    '₀': '0', '₁': '1', '₂': '2', '₃': '3', '₄': '4',
    '₅': '5', '₆': '6', '₇': '7', '₈': '8', '₉': '9',
    'ₙ': 'n', 'ₐ': 'a', 'ₑ': 'e', 'ₓ': 'x', 'ᵢ': 'i',
}

SYM2LATEX = {
    '⊆': r'\subseteq', '⊇': r'\supseteq', '⊂': r'\subset', '⊃': r'\supset',
    '⫋': r'\subsetneqq', '⫌': r'\supsetneqq', '⊊': r'\subsetneq',
    '∅': r'\varnothing', '∈': r'\in', '∉': r'\notin', '∋': r'\ni',
    '∩': r'\cap', '∪': r'\cup', '∁': r'\complement',
    '≤': r'\leq', '≥': r'\geq', '≠': r'\neq', '≈': r'\approx',
    '≡': r'\equiv', '≅': r'\cong', '∼': r'\sim', '≪': r'\ll', '≫': r'\gg',
    '±': r'\pm', '∓': r'\mp', '×': r'\times', '÷': r'\div', '⋅': r'\cdot',
    '∞': r'\infty', '∑': r'\sum', '∏': r'\prod', '∫': r'\int',
    '⇒': r'\Rightarrow', '⇐': r'\Leftarrow', '⇔': r'\Leftrightarrow',
    '→': r'\to', '←': r'\leftarrow', '↔': r'\leftrightarrow',
    '∀': r'\forall', '∃': r'\exists', '∴': r'\therefore', '∵': r'\because',
    '∠': r'\angle', '⊥': r'\perp', '∥': r'\parallel', '△': r'\triangle',
    'α': r'\alpha', 'β': r'\beta', 'γ': r'\gamma', 'δ': r'\delta',
    'ε': r'\epsilon', 'θ': r'\theta', 'λ': r'\lambda', 'μ': r'\mu',
    'π': r'\pi', 'ρ': r'\rho', 'σ': r'\sigma', 'τ': r'\tau',
    'φ': r'\phi', 'ω': r'\omega',
    'Γ': r'\Gamma', 'Δ': r'\Delta', 'Θ': r'\Theta', 'Λ': r'\Lambda',
    'Π': r'\Pi', 'Σ': r'\Sigma', 'Φ': r'\Phi', 'Ψ': r'\Psi', 'Ω': r'\Omega',
    'ℕ': r'\mathbb{N}', 'ℤ': r'\mathbb{Z}', 'ℚ': r'\mathbb{Q}',
    'ℝ': r'\mathbb{R}', 'ℂ': r'\mathbb{C}',
    '∣': r'\mid', '°': r'^\circ', '′': "'",
}

# 判定为"数学表达式"的强特征
REL_OPS = set('⊆⊇⊂⊃⫋⫌⊊⊋∈∉∋≠≤≥≈≡≅∼=<>⇒⇐⇔→←↔∀∃')
STRUCT_CHARS = set('√∑∏∫∞')

# 注意：切分必须用「连续中文块」(带 +)，且用捕获组 split。
# 早期版本用单字符正则 + findall 交错，split 过滤空串后两者数量对不齐，
# 导致中文块错位（"即"被塞进公式、"且"被留在外面）。
CJK_CHUNK = r'([\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]+)'
CJK_RE = re.compile(CJK_CHUNK)

# 占位符：保护"原始文本里的花括号"，避免与 LaTeX 生成的花括号混淆
LB, RB = '\x01', '\x02'


# ---------------------------------------------------------------- 切分

def _unbalanced(s):
    """返回括号净差值：>0 表示有未闭合的左括号"""
    return s.count('(') - s.count(')') + s.count('{') - s.count('}') \
        + s.count('[') - s.count(']')


def split_math_segments(text):
    """
    中文混排 -> [(is_math, content), ...]

    策略：
    1. 按中文切块（非中文块是"疑似公式"候选）
    2. 若某块括号未闭合（如 `{x | x∈A 且 x∈B}` 被中文"且"切开），
       吞并后续块直到闭合 —— 中文用 \\text{} 保留在公式内
    3. 候选块含强数学特征才判定为公式
    """
    # 带捕获组的 split：结果中奇数索引是中文块，偶数是非中文块，位置天然对齐
    parts = re.split(CJK_CHUNK, text)
    raw = [(idx % 2 == 1, p) for idx, p in enumerate(parts) if p]

    # 合并未闭合块
    merged = []
    buf, bufsrc = '', []
    for is_cjk, content in raw:
        if bufsrc:
            # 正在吞并
            buf += content
            bufsrc.append(content)
            if _unbalanced(buf) <= 0:
                merged.append((False, buf)); buf, bufsrc = '', []
            continue
        if not is_cjk and _unbalanced(content) > 0:
            buf, bufsrc = content, [content]
            continue
        merged.append((is_cjk, content))
    if buf:
        merged.append((False, buf))

    # 判定公式
    out = []
    for is_cjk, content in merged:
        if not is_cjk and _looks_like_math(content):
            out.append((True, content))
        else:
            out.append((False, content))
    return _merge_same(out)


def _looks_like_math(s):
    s = s.strip()
    if not s:
        return False
    if any(c in REL_OPS for c in s):
        return True
    if any(c in UNI_SUP or c in UNI_SUB for c in s):
        return True
    if re.search(r'[0-9A-Za-z]\^\s*[0-9A-Za-z{]', s):
        return True
    if any(c in STRUCT_CHARS for c in s):
        return True
    if re.search(r'[0-9A-Za-z)]/\s*[0-9A-Za-z(]', s):
        return True
    return False


def _merge_same(parts):
    out = []
    for flag, content in parts:
        if out and out[-1][0] == flag:
            out[-1] = (flag, out[-1][1] + content)
        else:
            out.append((flag, content))
    return [(f, c) for f, c in out if c]


# ---------------------------------------------------------------- 转换

def _match_paren(s, i):
    depth = 0
    for j in range(i, len(s)):
        if s[j] == '(':
            depth += 1
        elif s[j] == ')':
            depth -= 1
            if depth == 0:
                return j
    return -1


def convert_sqrt(s):
    """√(...) -> \\sqrt{...}"""
    out, i = [], 0
    while i < len(s):
        if s[i] == '√':
            j = i + 1
            if j < len(s) and s[j] == '(':
                k = _match_paren(s, j)
                if k > 0:
                    out.append(r'\sqrt{' + convert_sqrt(s[j + 1:k]) + '}')
                    i = k + 1
                    continue
            m = re.match(r'[0-9A-Za-z]+', s[j:])
            if m:
                out.append(r'\sqrt{' + m.group(0) + '}')
                i = j + len(m.group(0))
                continue
            out.append(r'\sqrt{}')
            i += 1
        else:
            out.append(s[i]); i += 1
    return ''.join(out)


def convert_frac(s):
    """a/b -> \\frac{a}{b}，从内到外处理"""
    def repl(m):
        return r'\frac{' + m.group(1) + '}{' + m.group(2) + '}'

    for _ in range(4):  # 处理嵌套
        before = s
        s = re.sub(r'\(([^()]+)\)/\s*\(([^()]+)\)', repl, s)
        s = re.sub(r'\(([^()]+)\)/\s*([0-9A-Za-z]+)', repl, s)
        s = re.sub(r'([0-9A-Za-z]+)/\s*\(([^()]+)\)', repl, s)
        s = re.sub(r'([0-9A-Za-z]+)/\s*([0-9A-Za-z]+)', repl, s)
        if s == before:
            break
    return s


def unicode_to_latex(text):
    """
    Unicode 伪公式 -> LaTeX

    关键：原始花括号先换成占位符 \\x01 \\x02，全部处理完再转回 \\{ \\}，
    避免 LaTeX 自己生成的花括号被误转义（早期版本的 bug）。
    """
    s = text.strip()
    if not s:
        return s

    # 0. 保护原始花括号
    s = s.replace('{', LB).replace('}', RB)

    # 1. 根号
    s = convert_sqrt(s)

    # 2. 分数
    s = convert_frac(s)

    # 3. Unicode 上下标
    out = []
    for ch in s:
        if ch in UNI_SUP:
            out.append('^{' + UNI_SUP[ch] + '}')
        elif ch in UNI_SUB:
            out.append('_{' + UNI_SUB[ch] + '}')
        else:
            out.append(ch)
    s = ''.join(out)

    # 4. 数学符号（长名优先，避免 ⊆ 被 ⊂ 抢先替换）
    for ch in sorted(SYM2LATEX, key=len, reverse=True):
        if ch in s:
            s = s.replace(ch, ' ' + SYM2LATEX[ch] + ' ')

    # 5. 公式内中文 -> \text{...}
    def cjk_repl(m):
        return r'\text{' + m.group(0) + '}'
    s = CJK_RE.sub(cjk_repl, s)

    # 6. 占位符还原为转义花括号
    s = s.replace(LB, r'\{').replace(RB, r'\}')

    # 7. 压缩空格
    s = re.sub(r'[ \t]+', ' ', s).strip()
    return s


# ---------------------------------------------------------------- 扫描 / 应用

def scan_docx(path):
    from docx import Document
    from docx.oxml.ns import qn

    doc = Document(path)
    findings = []
    for pi, para in enumerate(doc.paragraphs):
        if para._p.findall('.//' + qn('m:oMath')):
            continue
        full = para.text
        if not full or not any(c in SYM2LATEX or c in UNI_SUP or c in UNI_SUB
                               for c in full):
            continue
        for is_math, seg in split_math_segments(full):
            if not is_math:
                continue
            latex = unicode_to_latex(seg)
            level = 'high' if any(op in seg for op in REL_OPS) \
                or '\\sqrt' in latex or '\\frac' in latex else 'mid'
            findings.append({'para': pi, 'raw': seg, 'latex': latex,
                             'level': level})
    return doc, findings


def _latex_ok(latex):
    """校验 LaTeX 能被解析成有效 OMML"""
    try:
        from latex2omml import latex_to_omml
        xml = latex_to_omml(latex).decode('utf-8')
        return '<m:oMath' in xml and '&#' not in xml
    except Exception:
        return False


def _make_run(text, rpr_copy):
    """构造 w:r，继承原 run 的格式（加粗/颜色/字号等）"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from copy import deepcopy

    r = OxmlElement('w:r')
    if rpr_copy is not None:
        r.append(deepcopy(rpr_copy))
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t)
    return r


def apply_upgrade(path, min_level='mid', backup=False, verify=True,
                  outdir=None):
    """
    逐 run 替换伪公式为 oMath。

    关键：在 run 级别操作并复制原 rPr，**保留加粗/颜色/字号等格式**。
    （早期版本整段删除重建，会清空所有格式 —— 已修复）
    """
    from docx import Document
    from docx.oxml import parse_xml
    from docx.oxml.ns import qn
    from copy import deepcopy
    from latex2omml import latex_to_omml

    doc = Document(path)
    ok = fail = skipped = 0

    for para in doc.paragraphs:
        for run in list(para.runs):
            text = run.text
            if not text or not any(c in SYM2LATEX or c in UNI_SUP
                                   or c in UNI_SUB for c in text):
                continue
            segs = split_math_segments(text)
            if not any(is_m for is_m, _ in segs):
                continue
            # 置信度过滤
            if min_level == 'high':
                segs = [(m, c) for m, c in segs
                        if not m or any(op in c for op in REL_OPS)
                        or '\\sqrt' in unicode_to_latex(c)
                        or '\\frac' in unicode_to_latex(c)
                        or any(c2 in UNI_SUP or c2 in UNI_SUB for c2 in c)]
                if not any(m for m, _ in segs):
                    continue

            rpr = run._element.find(qn('w:rPr'))
            rpr_copy = deepcopy(rpr) if rpr is not None else None
            parent = run._element.getparent()
            idx = parent.index(run._element)

            new_els = []
            for is_math, content in segs:
                if not is_math:
                    new_els.append(_make_run(content, rpr_copy))
                    continue
                latex = unicode_to_latex(content)
                if verify and not _latex_ok(latex):
                    new_els.append(_make_run(content, rpr_copy))
                    skipped += 1
                    continue
                try:
                    new_els.append(parse_xml(latex_to_omml(latex)))
                    ok += 1
                except Exception as e:
                    new_els.append(_make_run(content, rpr_copy))
                    fail += 1
                    sys.stderr.write(f'  [失败] {content!r} -> {e}\n')

            parent.remove(run._element)
            for i, el in enumerate(new_els):
                parent.insert(idx + i, el)

    if ok == 0:
        return None, 0, fail, skipped

    if backup:
        shutil.copy2(path, path.replace('.docx', '.bak.docx'))

    name = os.path.basename(path)
    stem = name[:-5] if name.lower().endswith('.docx') else name
    if outdir:
        os.makedirs(outdir, exist_ok=True)
        out = os.path.join(outdir, stem + '.docx')
    else:
        out = path.replace('.docx', '_公式升级.docx')
    doc.save(out)
    return out, ok, fail, skipped


# ---------------------------------------------------------------- CLI

def main():
    ap = argparse.ArgumentParser(
        description='把 docx 里 Unicode 硬拼的伪公式升级为 Word 原生公式')
    ap.add_argument('target', help='docx 文件或目录')
    ap.add_argument('--apply', action='store_true', help='实际写入（默认只读）')
    ap.add_argument('--min-level', choices=['high', 'mid'], default='mid')
    ap.add_argument('--sample', type=int, default=0, help='打印 N 条转换抽样')
    ap.add_argument('--no-backup', action='store_true')
    ap.add_argument('--outdir', help='输出到指定目录（默认在原文件旁生成 '
                                     '_公式升级.docx）')
    args = ap.parse_args()

    if os.path.isdir(args.target):
        import glob
        files = sorted(glob.glob(os.path.join(args.target, '*.docx')))
    else:
        files = [args.target]

    tot_find = tot_ok = tot_fail = tot_skip = 0
    for f in files:
        name = os.path.basename(f)
        try:
            if args.apply:
                out, ok, fail, skip = apply_upgrade(
                    f, args.min_level, backup=not args.no_backup,
                    outdir=args.outdir)
                tot_ok += ok; tot_fail += fail; tot_skip += skip
                print(f'{name:<40} 升级 {ok:>3} 处，失败 {fail}，'
                      f'校验跳过 {skip}'
                      + (f' -> {os.path.basename(out)}' if out else ''))
            else:
                _, findings = scan_docx(f)
                hi = sum(1 for x in findings if x['level'] == 'high')
                mi = len(findings) - hi
                tot_find += len(findings)
                print(f'{name:<40} 发现 {len(findings):>3} 处 '
                      f'(高置信 {hi} / 中置信 {mi})')
                if args.sample:
                    for x in findings[:args.sample]:
                        print(f'      {x["raw"]!r}')
                        print(f'        -> {x["latex"]}')
        except Exception as e:
            print(f'{name:<40} 错误: {e}')

    print()
    if args.apply:
        print(f'合计：升级 {tot_ok}，失败 {tot_fail}，校验跳过 {tot_skip}')
    else:
        print(f'合计发现 {tot_find} 处可升级项（只读模式，未修改任何文件）')
        print('加 --apply 写入；默认自动备份 .bak.docx 并输出 _公式升级.docx')


if __name__ == '__main__':
    main()
