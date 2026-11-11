# -*- coding: utf-8 -*-
"""成品 docx 的公式符号污染体检 / 修复（只动 m:t 里的字符，不碰任何结构）。

背景
----
latex2omml 依赖的 latex2mathml 上游符号表有三处码点笔误（v1.1.5 起引擎
导入时自动打补丁）：\\sim→~、\\ast→*、\\cdot→·。补丁只影响**新生成**的
文档；在此之前转换出来的成品 docx 里，错误的 ASCII 字符已经固化在 <m:t> 中，
需要单独扫出来修。

三档处置（不要一刀切）
----------------------
  U+00B7 ·  -> U+22C5 ⋅   默认修。点乘 / 单位乘，是 \\cdot 笔误的真实产物。
  U+007E ~  -> U+223C ∼   默认修。相似、服从分布，是 \\sim 笔误的真实产物。
  U+002A *  -> U+2217 ∗   默认**不修**，需 --stars。抽样看几乎都是 n∈N*
                          这类原稿自带的上标星号，不是 \\ast 笔误；改成
                          二元运算符 ∗ 反而会带来多余间距。

安全约定
--------
  默认只读扫描；只有显式 --fix 才写文件，且默认写到副本
  <原名>_符号修正.docx，原件永远不动。--inplace 才原地替换并留 .bak。
  替换用字符串级正则限定在 <m:t>...</m:t> 内，不经 XML 解析器重排，
  不会改变命名空间声明、属性顺序或 Word 兼容性。

用法
----
    python fix_docx_symbols.py 文件.docx                    # 只读体检
    python fix_docx_symbols.py 目录/ --recursive            # 批量只读体检
    python fix_docx_symbols.py 文件.docx --fix              # 修到副本
    python fix_docx_symbols.py 文件.docx --fix --stars      # 连星号一起修
    python fix_docx_symbols.py 文件.docx --fix --inplace    # 原地 + .bak
    python fix_docx_symbols.py --selftest                  # 回归断言
    python fix_docx_symbols.py 文件.docx --json             # 机器可读

退出码：0 正常 / 1 有命中但只读 / 2 参数或文件错误。
"""
import argparse
import json
import os
import re
import shutil
import sys
import zipfile

M_T = re.compile(r'(<m:t(?:\s[^>]*)?>)(.*?)(</m:t>)', re.S)

# 错误字符 -> (正确字符, 说明, 是否默认修)
FIXES = {
    '\u00b7': ('\u22c5', 'U+00B7 文本中点 -> U+22C5 数学点乘 (\\cdot)', True),
    '\u007e': ('\u223c', 'U+007E ASCII 波浪 -> U+223C 数学相似 (\\sim)', True),
    '\u002a': ('\u2217', 'U+002A ASCII 星号 -> U+2217 数学星号 (\\ast)', False),
}


def active_fixes(include_stars=False):
    return {k: v for k, v in FIXES.items() if v[2] or (include_stars and k == '\u002a')}


def math_texts(xml):
    """产出 (整段 m:t 标签, 起始标签, 文本内容)，只取公式内的文本节点。"""
    for m in M_T.finditer(xml):
        yield m.group(0), m.group(1), m.group(2)


def inspect_file(path, fixes):
    """只读统计，返回 {'counts': {...}, 'samples': [...]}"""
    counts = {k: 0 for k in fixes}
    samples = []
    with zipfile.ZipFile(path) as z:
        if 'word/document.xml' not in z.namelist():
            return None
        xml = z.read('word/document.xml').decode('utf-8')
    for _full, _open, text in math_texts(xml):
        for ch in fixes:
            n = text.count(ch)
            if n:
                counts[ch] += n
                if len(samples) < 5:
                    samples.append(text[:80])
    return {'counts': counts, 'samples': samples, 'total': sum(counts.values())}


def rewrite_xml(xml, fixes):
    """只替换 m:t 内的字符，返回 (新 xml, 替换次数)。结构字节不变。"""
    total = [0]

    def repl(m):
        body = m.group(2)
        new = body
        for ch, (good, _desc, _default) in fixes.items():
            if ch in new:
                total[0] += new.count(ch)
                new = new.replace(ch, good)
        return m.group(1) + new + m.group(3) if new != body else m.group(0)

    return M_T.sub(repl, xml), total[0]


def write_docx(src, dst, fixes):
    """复制 zip 全部条目，仅替换 word/document.xml。返回替换次数。"""
    with zipfile.ZipFile(src) as zin:
        xml = zin.read('word/document.xml').decode('utf-8')
        new_xml, n = rewrite_xml(xml, fixes)
        if not n:
            return 0
        items = [(i, zin.read(i.filename)) for i in zin.infolist()
                 if i.filename != 'word/document.xml']
        with zipfile.ZipFile(dst, 'w') as zout:
            for info, data in items:
                zout.writestr(info, data)
            zout.writestr('word/document.xml', new_xml.encode('utf-8'),
                          compress_type=zipfile.ZIP_DEFLATED)
    return n


def collect(paths, recursive):
    out = []
    for p in paths:
        if os.path.isdir(p):
            if recursive:
                for dp, _dns, fns in os.walk(p):
                    for fn in fns:
                        if fn.lower().endswith('.docx') and not fn.startswith('~$'):
                            out.append(os.path.join(dp, fn))
            else:
                for fn in os.listdir(p):
                    if fn.lower().endswith('.docx') and not fn.startswith('~$'):
                        out.append(os.path.join(p, fn))
        elif p.lower().endswith('.docx'):
            out.append(p)
        else:
            print('跳过（不是 docx）：%s' % p)
    return out


def selftest():
    """回归：替换只在 m:t 内发生；正文 run、属性、结构都不受影响。"""
    fixes = active_fixes(include_stars=True)
    xml = (
        '<w:document xmlns:w="W" xmlns:m="M">'
        '<w:p><w:r><w:t>正文里有 · 和 ~ 和 *，不该被改</w:t></w:r>'
        '<m:oMath><m:r><m:t>a·b~c*d</m:t></m:r>'
        '<m:r><m:t xml:space="preserve"> · </m:t></m:r></m:oMath></w:p></w:document>'
    )
    new, n = rewrite_xml(xml, fixes)
    checks = [
        ('替换次数=4', n == 4),
        ('正文未动', '<w:t>正文里有 · 和 ~ 和 *，不该被改</w:t>' in new),
        ('点乘已改', '<m:t>a\u22c5b\u223cc\u2217d</m:t>' in new),
        ('属性保留', '<m:t xml:space="preserve"> \u22c5 </m:t>' in new),
        ('结构未变', new.count('<m:oMath>') == 1 and new.count('</w:document>') == 1),
    ]
    # 默认不含星号
    dflt, n2 = rewrite_xml(xml, active_fixes(False))
    checks.append(('默认不动星号', n2 == 3 and '<m:t>a\u22c5b\u223cc*d</m:t>' in dflt))
    for name, ok in checks:
        print('  %s %s' % ('PASS' if ok else 'FAIL', name))
    return 0 if all(ok for _n, ok in checks) else 1


def main():
    ap = argparse.ArgumentParser(description='成品 docx 公式符号体检 / 修复')
    ap.add_argument('paths', nargs='*', help='docx 文件或目录')
    ap.add_argument('--recursive', '-r', action='store_true', help='目录递归')
    ap.add_argument('--fix', action='store_true', help='写文件（默认只读）')
    ap.add_argument('--stars', action='store_true', help='连 U+002A 星号一起修（默认不动）')
    ap.add_argument('--inplace', action='store_true', help='原地替换并留 .bak（需 --fix）')
    ap.add_argument('--output', help='指定输出路径（仅单文件 + --fix 有效）')
    ap.add_argument('--json', action='store_true', help='JSON 输出')
    ap.add_argument('--selftest', action='store_true', help='回归断言')
    args = ap.parse_args()

    if args.selftest:
        print('fix_docx_symbols 自检')
        return selftest()
    if not args.paths:
        ap.error('需要至少一个 docx 路径，或用 --selftest')
    if args.inplace and not args.fix:
        ap.error('--inplace 需要配合 --fix')
    if args.output and (not args.fix or len(args.paths) != 1):
        ap.error('--output 只在单文件 + --fix 时可用')

    fixes = active_fixes(args.stars)
    files = collect(args.paths, args.recursive)
    if not files:
        print('没有找到 docx')
        return 2

    rows = []
    for p in files:
        try:
            r = inspect_file(p, fixes)
            if r and r['total']:
                rows.append({'file': p, **r})
        except (zipfile.BadZipFile, OSError, UnicodeDecodeError) as exc:
            if not args.json:
                print('  跳过（读不了）：%s  %s' % (p, exc))

    if args.json:
        print(json.dumps({'fixes': {('U+%04X' % ord(k)): v[1] for k, v in fixes.items()},
                          'files': len(files), 'hits': rows}, ensure_ascii=False, indent=2))
        return 0

    print('公式符号污染体检（%d 个 docx，%d 个命中）' % (len(files), len(rows)))
    print('=' * 66)
    print('启用修正：')
    for ch, (good, desc, default) in fixes.items():
        print('  %s -> %s   %s' % (ch, good, desc))
    print('  注：U+002A 星号默认不动（多为 n∈N* 原稿写法）；确需修加 --stars')
    print('-' * 66)
    for r in rows:
        detail = ', '.join('%s×%d' % ('U+%04X' % ord(k), v)
                           for k, v in r['counts'].items() if v)
        print('  %-4d %s\n       %s' % (r['total'], r['file'], detail))
        for s in r['samples'][:2]:
            print('       例: %r' % s)
    print('=' * 66)

    if not args.fix:
        if rows:
            print('只读体检完成，未修改任何文件。要修复请加 --fix（默认输出到副本）。')
        else:
            print('未发现污染。')
        return 0

    changed = 0
    for r in rows:
        src = r['file']
        stem, ext = os.path.splitext(src)
        if args.output:
            dst = args.output
        elif args.inplace:
            dst = src + '.new'
        else:
            dst = stem + '_符号修正' + ext
        try:
            n = write_docx(src, dst, fixes)
            if args.inplace:
                shutil.copy2(src, stem + '.bak' + ext)
                os.replace(dst, src)
                dst = src
            print('  ✓ %s  （%d 处）' % (dst, n))
            changed += n
        except (OSError, zipfile.BadZipFile) as exc:
            print('  ✗ 写失败 %s：%s（原件未动）' % (src, exc))
    print('完成，共修正 %d 处。' % changed)
    return 0


if __name__ == '__main__':
    sys.exit(main())
