# -*- coding: utf-8 -*-
"""符号库体检：上游码点表勘误 + 文档/引擎一致性 + 字形覆盖。

三类问题分开报，别混为一谈：
  A 类（必坏）  上游 SYMBOLS 里码点写错 —— Word 会显示成另一个字符，且不报错
  B 类（待查）  references/symbols.md 写的符号与引擎实际输出对不上
  C 类（信息）  上游表里 Cambria Math 无字形的码点 —— 用到才会丢字，不是错误

用法：
    python scripts/scan_symbol_table.py               # 体检报告
    python scripts/scan_symbol_table.py --all         # 列出全部 C 类明细
    python scripts/scan_symbol_table.py --selftest    # 回归断言（供 CI / 自检用）
    python scripts/scan_symbol_table.py --json        # 机器可读输出

退出码：0 = A/B 类均为 0；1 = 有 A/B 类问题；2 = 依赖缺失跑不起来。
"""
import argparse
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
sys.path.insert(0, HERE)

try:
    import latex2omml as L
    from lxml import etree
    from latex2mathml import symbols_parser as SP
except ImportError as exc:                                     # pragma: no cover
    print('依赖缺失：%s' % exc)
    print('安装：pip install lxml latex2mathml python-docx')
    sys.exit(2)

M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
HEX5 = re.compile(r'[0-9A-Fa-f]{5}\Z')

# ---------------------------------------------------------------- A 类参照表
# 常用 LaTeX 命令 -> 标准 Unicode 码点。只收「有公认唯一答案」的命令，
# 有异议的（\colon→:、\vert→|、\less→<）一律不收，避免误报。
CANONICAL = {
    r'\sim': 0x223C, r'\ast': 0x2217, r'\cdot': 0x22C5,
    r'\times': 0x00D7, r'\div': 0x00F7, r'\pm': 0x00B1,
    r'\mp': 0x2213, r'\circ': 0x2218, r'\oplus': 0x2295,
    r'\otimes': 0x2297, r'\leq': 0x2264, r'\geq': 0x2265,
    r'\neq': 0x2260, r'\approx': 0x2248, r'\equiv': 0x2261,
    r'\simeq': 0x2243, r'\cong': 0x2245, r'\propto': 0x221D,
    r'\in': 0x2208, r'\notin': 0x2209, r'\subset': 0x2282,
    r'\subseteq': 0x2286, r'\supset': 0x2283, r'\supseteq': 0x2287,
    r'\cup': 0x222A, r'\cap': 0x2229, r'\varnothing': 0x2205,
    r'\infty': 0x221E, r'\partial': 0x2202, r'\nabla': 0x2207,
    r'\sum': 0x2211, r'\prod': 0x220F, r'\int': 0x222B,
    r'\oint': 0x222E, r'\rightarrow': 0x2192, r'\Rightarrow': 0x21D2,
    r'\Longrightarrow': 0x27F9, r'\leftrightarrow': 0x2194,
    r'\Leftrightarrow': 0x21D4, r'\backsim': 0x223D,
    r'\angle': 0x2220, r'\perp': 0x22A5, r'\parallel': 0x2225,
    r'\because': 0x2235, r'\therefore': 0x2234,
    r'\ldots': 0x2026, r'\cdots': 0x22EF,
    r'\alpha': 0x03B1, r'\beta': 0x03B2, r'\pi': 0x03C0,
    r'\Omega': 0x03A9, r'\aleph': 0x2135, r'\ell': 0x2113,
}

# 经 premap 改写成结构型表达式（上标 / 组合）的命令，不是单字符映射。
# A 类判据是「最终 OMML 文本 == 一个标准字符」，对它们不成立，须排除：
#   \degree -> ^{\circ}，输出带 m:sSup 的 ∘ U+2218（LaTeX 度数标准写法）
#   \degreeCelsius / \celsius -> ^{\circ}\mathrm{C}
# 注意：这不是笔误；直接写 ° U+00B0 只会落在基线上，结构反而是错的。
STRUCTURAL_CMDS = {r'\degree', r'\degreeCelsius', r'\celsius'}

# 这些命令落到 ASCII 是 LaTeX 本来的语义，属于正常映射，不算笔误
LEGIT_ASCII = {
    r'\sptilde', r'\sphat', r'\backslash', r'\colon', r'\mathcolon',
    r'\vert', r'\lvert', r'\rvert', r'\less', r'\lt',
    r'\greater', r'\gt', r'\equal', r'\plus', r'\comma',
    r'\period', r'\ldotp', r'\semicolon', r'\question', r'\atsign',
    r'\exclam', r'\octothorpe', r'\mathdollar', r'\percent',
    r'\ampersand', r'\binampersand', r'\And', r'\lparen', r'\rparen',
    r'\lbrack', r'\rbrack', r'\lbrace', r'\rbrace', r'\mathslash',
    r'\slash', r'\_', r'\{', r'\}', r'\#', r'\$', r'\%', r'\&',
}


def scan_upstream():
    """返回 (A类问题, ASCII退化命令数, 已补偿的命令, 补偿明细)

    判据是**最终 OMML 文本**，不是上游表的原始值：上游表写错但引擎
    （_MATHML_CHAR_FIX / _patch_upstream_symbols）已经补偿的，只算「已补偿」，
    不算故障。否则像 \\perp（上游 27C2 -> 引擎已改 22A5）会一直误报。
    """
    issues, ascii_cmds, compensated = [], [], []
    for cmd, val in SP.SYMBOLS.items():
        if not (isinstance(val, str) and HEX5.match(val)):
            continue
        cp = int(val, 16)
        if cp < 0x80 and cmd.startswith('\\') and cmd not in LEGIT_ASCII:
            ascii_cmds.append((cmd, cp))

    patches = set(getattr(L, 'PATCHED_SYMBOLS', []))
    char_fixes = set(getattr(L, '_MATHML_CHAR_FIX', {}))
    for cmd, want in CANONICAL.items():
        if cmd in STRUCTURAL_CMDS:
            continue
        raw = SP.SYMBOLS.get(cmd)
        if raw is None or not HEX5.match(raw):
            continue
        want_char = chr(want)
        try:
            final = omml_text(cmd)
        except Exception as exc:                            # noqa: BLE001
            issues.append({'cmd': cmd, 'got': 'ERROR: %s' % exc,
                           'want': want_char, 'raw': '—'})
            continue
        if final != want_char:
            issues.append({'cmd': cmd, 'got': final, 'want': want_char,
                           'raw': 'U+%04X' % int(raw, 16), 'fixed_by_patch': False})
        elif int(raw, 16) != want:
            how = ('上游表补丁' if cmd in patches else
                   '字符纠错表' if chr(int(raw, 16)) in char_fixes else '其他补偿')
            compensated.append({'cmd': cmd, 'raw': int(raw, 16),
                                'want': want, 'how': how})
    return issues, len(ascii_cmds), patches, compensated


def omml_text(latex):
    s = L.latex_to_omml(latex, display='inline')
    s = s.decode('utf8') if isinstance(s, bytes) else s
    return ''.join(t.text or '' for t in etree.fromstring(s.encode('utf8')).iter(M + 't'))


def doc_pairs():
    """从 references/symbols.md 的表格里抽出 (符号, LaTeX命令) 对"""
    path = os.path.join(SKILL, 'references', 'symbols.md')
    if not os.path.exists(path):
        return []
    out = []
    for line in open(path, encoding='utf-8'):
        if not line.strip().startswith('|'):
            continue
        cells = [c.strip() for c in line.strip().strip('|').split('|')]
        for a, b in zip(cells[0::2], cells[1::2]):
            ma = re.fullmatch(r'`([^`]+)`', a)
            mb = re.fullmatch(r'`([^`]+)`', b)
            if mb and len(a) == 1 and ord(a) > 0x2000:
                out.append((a, mb.group(1)))
            elif ma and len(b) == 1 and ord(b) > 0x2000:
                out.append((b, ma.group(1)))
    return out


SKIP_CMD = ('^', r'\text', r'\frac', r'\sqrt', r'\mathbf', r'\left',
            r'\begin', r'\mathrm', '{', '}')


def scan_doc():
    """返回 (B类问题列表, 校验过的对数)"""
    bad, checked = [], 0
    for ch, cmd in doc_pairs():
        if any(x in cmd for x in SKIP_CMD):
            continue
        checked += 1
        try:
            got = omml_text(cmd)
        except Exception as exc:                            # noqa: BLE001
            bad.append({'char': ch, 'code': ord(ch), 'cmd': cmd,
                        'got': 'ERROR: %s' % exc})
            continue
        if got != ch:
            bad.append({'char': ch, 'code': ord(ch), 'cmd': cmd, 'got': got})
    return bad, checked


def cambria_cmap():
    """取 Cambria Math 的 cmap；非 Windows / 缺 fontTools 时返回 None"""
    path = os.path.join(os.environ.get('WINDIR', r'C:\Windows'),
                        'Fonts', 'cambria.ttc')
    if not os.path.exists(path):
        return None
    try:
        from fontTools.ttLib import TTCollection
        fonts = TTCollection(path).fonts
        for f in fonts:
            if (f['name'].getDebugName(1) or '').lower().endswith('math'):
                return f.getBestCmap()
        return fonts[0].getBestCmap() if fonts else None
    except Exception:                                      # noqa: BLE001
        return None


def scan_font():
    """返回 (无字形条目数, 明细列表, 是否可用)"""
    cmap = cambria_cmap()
    if cmap is None:
        return 0, [], False
    missing = {}
    for cmd, val in SP.SYMBOLS.items():
        if not (isinstance(val, str) and HEX5.match(val)):
            continue
        cp = int(val, 16)
        if cp > 0x7F and cp not in cmap:
            missing.setdefault(cp, []).append(cmd)
    rows = [{'code': cp, 'char': chr(cp), 'cmds': cmds[:4]}
            for cp, cmds in sorted(missing.items())]
    return len(rows), rows, True


def main():
    ap = argparse.ArgumentParser(description='latex-omml 符号库体检')
    ap.add_argument('--all', action='store_true', help='列出全部 C 类明细')
    ap.add_argument('--json', action='store_true', help='JSON 输出')
    ap.add_argument('--selftest', action='store_true', help='回归断言')
    args = ap.parse_args()

    a, n_ascii, patches, compensated = scan_upstream()
    b, checked = scan_doc()
    n_font, font_rows, font_ok = scan_font()

    report = {
        'upstream_fixes': sorted(patches),
        'compensated': compensated,
        'ascii_degenerate': n_ascii,
        'A_issues': a,
        'B_checked': checked,
        'B_issues': b,
        'C_font_available': font_ok,
        'C_missing_glyphs': n_font,
        'C_detail': font_rows if (args.all or args.json) else font_rows[:10],
        'structural_cmds': sorted(STRUCTURAL_CMDS),
    }
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print('latex-omml 符号库体检')
        print('=' * 62)
        print('上游表条目 %d，退化到 ASCII 的反斜杠命令 %d 条（多数是字母/标点，正常）'
              % (len(SP.SYMBOLS), n_ascii))
        print('引擎已就地修正：%s' % (', '.join(sorted(patches)) or '（无，上游已是正确值）'))
        print('结构型命令（经 premap 改写，不适用单字符判据）：%s'
              % ', '.join(sorted(STRUCTURAL_CMDS)))
        if compensated:
            print('上游写错但已被补偿：%d 处' % len(compensated))
            for x in compensated:
                print('   %-14s U+%04X -> U+%04X  [%s]'
                      % (x['cmd'], x['raw'], x['want'], x['how']))
        print('A 类 · 最终 OMML 与标准码点不符：%d' % len(a))
        for x in a:
            got = x['got']
            got_desc = (' '.join('U+%04X' % ord(c) for c in got)
                        if isinstance(got, str) and not got.startswith('ERROR') else repr(got))
            print('   %-14s 实际 %r (%s) | 应为 %s | 上游原始 %s'
                  % (x['cmd'], got, got_desc, x['want'], x.get('raw', '—')))
        print('B 类 · 文档与引擎不一致：%d（共核对 %d 对）' % (len(b), checked))
        for x in b:
            print('   文档 %s U+%04X | %-14s -> 实际 %r' %
                  (x['char'], x['code'], x['cmd'], x['got']))
        if font_ok:
            print('C 类 · 上游表中 Cambria Math 无字形的码点：%d（信息项，用到才丢字）'
                  % n_font)
            for x in font_rows[:10]:
                print('   U+%04X %-3s  %s' % (x['code'], x['char'], ', '.join(x['cmds'])))
            if n_font > 10 and not args.all:
                print('   … 其余 %d 条用 --all 查看' % (n_font - 10))
        else:
            print('C 类 · 跳过（非 Windows 或未装 fontTools，不影响 A/B 结论）')
        print('=' * 62)

    if args.selftest:
        # 判据是「最终 OMML 输出」而不是上游表原始值：\perp 之类由字符纠错表
        # 兜底的命令，上游码点本来就不等于 CANONICAL，硬比原始值会误报。
        errors = []
        if a:
            errors.append('A 类仍有 %d 处不一致' % len(a))
        if b:
            errors.append('B 类仍有 %d 处文档不一致' % len(b))
        for cmd in getattr(L, 'PATCHED_SYMBOLS', []):
            if cmd not in patches:
                errors.append('%s 的上游表补丁丢失' % cmd)
        if not patches:
            errors.append('未检测到任何上游表补丁（补丁逻辑可能被整段删掉）')
        if errors:
            print('SELFTEST FAIL: ' + '; '.join(errors))
            return 1
        print('SELFTEST PASS: %d 条标准映射、%d 对文档条目全部一致；'
              '上游补丁 %d 条生效（%s）'
              % (len(CANONICAL), checked, len(patches), ', '.join(sorted(patches))))
        return 0
    return 1 if (a or b) else 0


if __name__ == '__main__':
    sys.exit(main())
