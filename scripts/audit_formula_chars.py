# -*- coding: utf-8 -*-
"""统一公式字符体检：docx（OMML / OLE / 图片）+ PDF（文字层 / 私有区）。

为什么需要它
------------
`fix_docx_symbols.py` 只看 `<m:t>`（OMML 原生公式里的文本节点）。但真实资料里
公式有多种形态：MathType OLE 对象、图片公式、Unicode 硬拼的伪公式，还有 PDF
文字层里的 Symbol/Euclid 字体私有区码点。只扫 m:t 会得出"没问题"的假结论。
本脚本把所有形态一起盘，并**按处置级别分开报**：

  ERROR  必修  —— 公式内 U+00B7 `·`、U+007E `~`、U+FFFD 替换符、私有区码点
  WARN   可选  —— 公式内全角标点、Unicode 伪上标（属结构问题，换字符解决不了）
  INFO   不动  —— `*` 星号（多为 n∈N*）、`□` 方框（多为填空占位）、
                  正文里的题源间隔号 `·`（如「2021·北京」）、OLE/图片公式数、
                  PDF 扫描件、EuclidMath 私有码点（不属于 Adobe Symbol 编码）

铁律
----
- 星号 `*` 永不自动修（用户 2026-11-11 明确确认）。
- 正文 `w:t` 里的 `·` 是中文题源标注，绝不改；修复只作用于 `<m:t>`。
- 默认只读；`--fix` 也只把 docx 修到副本，PDF 不做任何写入。

用法
----
    python scripts/audit_formula_chars.py 路径... -r
    python scripts/audit_formula_chars.py 路径... -r --json
    python scripts/audit_formula_chars.py 路径... -r --fix       # docx 修到副本
    python scripts/audit_formula_chars.py 路径... --pdf-full     # PDF 读全书
    python scripts/audit_formula_chars.py --selftest

退出码：0 无 ERROR；1 存在 ERROR（只读时）；2 参数/依赖错误。
"""
import argparse
import contextlib
import io
import json
import os
import re
import sys
import zipfile
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

try:
    import fix_docx_symbols as F
except ImportError as exc:                                     # pragma: no cover
    print('缺少同目录的 fix_docx_symbols.py：%s' % exc)
    sys.exit(2)

try:
    import pdf2docx_math as P
    HAVE_SYM = True
except Exception:                                              # pragma: no cover
    P = None
    HAVE_SYM = False

M_T = re.compile(r'<m:t(?:\s[^>]*)?>(.*?)</m:t>', re.S)
W_T = re.compile(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', re.S)

# 公式 m:t 内：可选修（改了更规范，但不是"错字"）
WARN_M = {
    '，': '全角逗号', '（': '全角左括号', '）': '全角右括号',
    '；': '全角分号', '：': '全角冒号',
    '²': '伪上标²(未用 sSup)', '³': '伪上标³', 'ⁿ': '伪上标ⁿ', 'ᵐ': '伪上标ᵐ',
}
# 公式 m:t 内：明确不动
INFO_M = {
    '*': '星号(多为 n∈N*，用户确认不修)',
    '□': '方框(多为填空占位，不修)',
}
# 正文 w:t 内：只统计，不动
INFO_W = {'·': '题源间隔号(如 2021·北京，不是点乘)', '□': '填空方框'}


@contextlib.contextmanager
def _mute():
    """压住 MuPDF 底层 xref 报错（写 stdout/stderr 都有，Windows 尤甚）。"""
    saved = {}
    try:
        for fd in (1, 2):
            try:
                saved[fd] = os.dup(fd)
                null = os.open(os.devnull, os.O_WRONLY)
                os.dup2(null, fd)
                os.close(null)
            except OSError:
                pass
        with contextlib.redirect_stdout(io.StringIO()), \
             contextlib.redirect_stderr(io.StringIO()):
            yield
    finally:
        for fd, old in saved.items():
            try:
                os.dup2(old, fd)
                os.close(old)
            except OSError:
                pass


def collect(paths, recursive):
    out = []
    for p in paths:
        if os.path.isdir(p):
            if recursive:
                for dp, _dns, fns in os.walk(p):
                    for fn in fns:
                        if not fn.startswith('~$'):
                            out.append(os.path.join(dp, fn))
            else:
                for fn in os.listdir(p):
                    q = os.path.join(p, fn)
                    if os.path.isfile(q):
                        out.append(q)
        elif os.path.isfile(p):
            out.append(p)
        else:
            print('跳过（不存在）：%s' % p)
    return out


def audit_docx(path):
    """返回单份 docx 的报告 dict。"""
    r = {'file': path, 'kind': 'docx', 'error': Counter(), 'warn': Counter(),
         'info': Counter(), 'shapes': {}, 'samples': []}
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if 'word/document.xml' not in names:
            return None
        xml = z.read('word/document.xml').decode('utf-8', 'replace')
    # ERROR：复用 fix_docx_symbols 的判据（默认修 · 与 ~，不含星号）
    ins = F.inspect_file(path, F.active_fixes(False))
    if ins:
        for ch, n in ins['counts'].items():
            if n:
                r['error']['U+%04X %s' % (ord(ch), ch)] += n
                r['samples'] += ins['samples'][:2]
    mt = ''.join(M_T.findall(xml))
    wt = ''.join(W_T.findall(xml))
    for ch, label in WARN_M.items():
        n = mt.count(ch)
        if n:
            r['warn'][label] += n
    for ch, label in INFO_M.items():
        n = mt.count(ch)
        if n:
            r['info'][label] += n
    for ch, label in INFO_W.items():
        n = wt.count(ch)
        if n:
            r['info']['正文·' + label] += n
    if '\ufffd' in mt:
        r['error']['U+FFFD 替换符'] += mt.count('\ufffd')
    priv = sum(1 for c in mt if 0xE000 <= ord(c) <= 0xF8FF)
    if priv:
        r['error']['私有区码点(公式内)'] += priv
    r['shapes'] = {
        'oMath': xml.count('<m:oMath'),
        'OLE': xml.count('<w:object') + xml.count('<o:OLEObject'),
        'embeddings': sum(1 for n in names if n.startswith('word/embeddings/')),
        'drawings': xml.count('<w:drawing') + xml.count('<w:pict'),
    }
    return r


def sym_coverage(counter):
    """私有区码点 -> (已覆盖次数, 未覆盖明细)。"""
    if not HAVE_SYM:
        return 0, []
    covered, miss = 0, []
    for ch, n in counter.items():
        if (ord(ch) & 0xFF) in P.SYM:
            covered += n
        else:
            miss.append({'code': ord(ch), 'count': n})
    return covered, sorted(miss, key=lambda x: -x['count'])


def audit_pdf(path, max_pages=20, full=False):
    """返回单份 PDF 的报告 dict；缺 pymupdf 时返回带 error 字段的报告。"""
    r = {'file': path, 'kind': 'pdf', 'pages': 0, 'text_chars': 0,
         'private': Counter(), 'unknown': [], 'scanned': False,
         'info': Counter(), 'warn': Counter(), 'error': Counter()}
    try:
        import pymupdf
    except ImportError:
        r['note'] = '未装 pymupdf，跳过 PDF（pip install pymupdf）'
        return r
    try:
        with _mute():
            doc = pymupdf.open(path)
            r['pages'] = len(doc)
            n = len(doc) if full else min(max_pages, len(doc))
            text = []
            for i in range(n):
                try:
                    text.append(doc[i].get_text())
                except Exception:
                    pass
            doc.close()
        txt = ''.join(text)
        r['text_chars'] = len(txt)
        r['scanned'] = len(txt) < 50
        r['private'] = Counter(c for c in txt if 0xF000 <= ord(c) <= 0xF0FF)
        covered, miss = sym_coverage(r['private'])
        r['private_covered'] = covered
        r['unknown'] = miss
        if miss:
            r['warn']['SYM 未覆盖的私有码点'] = sum(x['count'] for x in miss)
        if '\ufffd' in txt:
            r['error']['U+FFFD 替换符'] = txt.count('\ufffd')
        if txt.count('·'):
            r['info']['正文·题源间隔号'] = txt.count('·')
    except Exception as exc:
        r['note'] = '读取失败：%s' % str(exc)[:70]
    return r


def to_plain(r):
    """把 Counter 转成可 JSON 序列化的普通 dict。"""
    d = dict(r)
    for k in ('error', 'warn', 'info', 'private'):
        if isinstance(d.get(k), Counter):
            d[k] = dict(d[k])
    return d


def summarize(rows):
    s = {'files': Counter(), 'error': Counter(), 'warn': Counter(),
         'info': Counter(), 'shapes': Counter(), 'pdf_private': Counter(),
         'pdf_unknown': Counter(), 'scanned': 0}
    for r in rows:
        s['files'][r['kind']] += 1
        for key in ('error', 'warn', 'info'):
            s[key].update(r.get(key, {}))
        if r['kind'] == 'docx':
            for k, v in r['shapes'].items():
                s['shapes'][k] += v
        else:
            if r.get('scanned'):
                s['scanned'] += 1
            s['pdf_private'].update({('U+%04X' % ord(c)): n
                                     for c, n in r['private'].items()})
            for x in r['unknown']:
                s['pdf_unknown']['U+%04X' % x['code']] += x['count']
    return s


def selftest():
    checks = []

    def ck(name, ok):
        checks.append((name, bool(ok)))

    ck('Symbol 0x7E 应为 ∼ 而非 ~', HAVE_SYM and P.SYM.get(0x7E) == '∼')
    if HAVE_SYM:
        P.UNMAPPED_PRIVATE.clear()
        ck('remap U+F07E -> ∼', P.remap_symbol('\uf07e') == '∼')
        ck('未知 Euclid 码点被丢弃', P.remap_symbol('\uf08c') == '')
        ck('未知 Euclid 码点被记录', 0xF08C in P.unmapped_private_report())
        ck('已知空值码点不算未知', P.remap_symbol('\uf0ff') == ''
           and 0xF0FF not in P.unmapped_private_report())
        cov, miss = sym_coverage(Counter({'\uf07e': 3, '\uf08c': 2}))
        ck('覆盖率计算：3 覆盖 / 1 种未覆盖', cov == 3 and len(miss) == 1
           and miss[0]['code'] == 0xF08C)
    ck('· 属 ERROR 判据', '\u00b7' in F.active_fixes(False))
    ck('* 不在默认修复集', '\u002a' not in F.active_fixes(False))
    ck('* 已列入 INFO 不修', '*' in INFO_M)
    ck('□ 已列入 INFO 不修', '□' in INFO_M)
    ck('全角逗号属 WARN', '，' in WARN_M)
    xml = ('<w:t>2021·北京·月考</w:t>'
           '<m:t>a·b，c²*d□</m:t>')
    mt = ''.join(M_T.findall(xml))
    wt = ''.join(W_T.findall(xml))
    ck('正文 · 被单独识别（不进公式）', wt.count('·') == 2 and mt.count('·') == 1)
    ck('公式内 · 被识别', mt.count('·') == 1)
    for name, ok in checks:
        print('  %s %s' % ('PASS' if ok else 'FAIL', name))
    bad = [n for n, ok in checks if not ok]
    if bad:
        print('SELFTEST FAIL: %s' % '; '.join(bad))
        return 1
    print('SELFTEST PASS: %d 项' % len(checks))
    return 0


def main():
    ap = argparse.ArgumentParser(description='统一公式字符体检（docx + PDF）')
    ap.add_argument('paths', nargs='*', help='文件或目录')
    ap.add_argument('-r', '--recursive', action='store_true', help='目录递归')
    ap.add_argument('--pdf-pages', type=int, default=20, help='每份 PDF 读前 N 页，默认 20')
    ap.add_argument('--pdf-full', action='store_true', help='PDF 读全书（慢）')
    ap.add_argument('--fix', action='store_true', help='把 ERROR 类 docx 修到副本')
    ap.add_argument('--json', action='store_true', help='JSON 输出')
    ap.add_argument('--selftest', action='store_true', help='回归断言')
    args = ap.parse_args()

    if args.selftest:
        print('audit_formula_chars 自检')
        return selftest()
    if not args.paths:
        ap.error('需要至少一个路径，或用 --selftest')

    files = collect(args.paths, args.recursive)
    rows, skipped = [], []
    for p in files:
        ext = os.path.splitext(p)[1].lower()
        try:
            if ext == '.docx':
                r = audit_docx(p)
            elif ext == '.pdf':
                r = audit_pdf(p, args.pdf_pages, args.pdf_full)
            else:
                skipped.append(p)
                continue
            if r:
                rows.append(r)
        except (zipfile.BadZipFile, OSError, UnicodeDecodeError) as exc:
            skipped.append('%s (%s)' % (p, str(exc)[:40]))

    s = summarize(rows)
    if args.json:
        print(json.dumps({'summary': {k: (dict(v) if isinstance(v, Counter) else v)
                                      for k, v in s.items()},
                          'files': [to_plain(r) for r in rows],
                          'skipped': skipped}, ensure_ascii=False, indent=2))
        return 0

    print('统一公式字符体检')
    print('=' * 70)
    print('文件：docx %d / pdf %d / 其它跳过 %d'
          % (s['files']['docx'], s['files']['pdf'], len(skipped)))
    print('docx 公式形态：%s' % dict(s['shapes']))
    print('-' * 70)
    for title, key, hint in [
        ('ERROR 必修（公式内错字符）', 'error', '→ --fix 可修到副本'),
        ('WARN 可选（排版习惯 / 结构问题）', 'warn', '→ 默认不改'),
        ('INFO 不动（合法用法）', 'info', '→ 明确排除，别误伤'),
    ]:
        print('%s  %s' % (title, hint))
        if s[key]:
            for k, v in s[key].most_common():
                print('   %-40s %d' % (k, v))
        else:
            print('   （无）')
    print('-' * 70)
    print('PDF：带私有区码点 %d 种 / %d 次；扫描件 %d 份'
          % (len(s['pdf_private']), sum(s['pdf_private'].values()), s['scanned']))
    if HAVE_SYM:
        cov = sum(r.get('private_covered', 0) for r in rows if r['kind'] == 'pdf')
        unk = sum(s['pdf_unknown'].values())
        total = cov + unk
        print('     SYM 覆盖 %d / %d 次（%.2f%%）；未覆盖 %d 次'
              % (cov, total, 100.0 * cov / total if total else 100.0, unk))
        if s['pdf_unknown']:
            print('     未覆盖明细：%s' % ', '.join(
                '%s×%d' % (k, v) for k, v in s['pdf_unknown'].most_common(12)))
            print('     ⚠️ 多为 EuclidMathOne/Two 私有字形，不属于 Adobe Symbol 编码；'
                  '转换时静默丢弃，需人工核对原页。')
    else:
        print('     （pdf2docx_math 不可用，跳过覆盖率计算）')
    print('=' * 70)

    err_files = [r for r in rows if r['kind'] == 'docx' and r['error']]
    if args.fix:
        print('修复（只写副本，原件不动）：')
        for r in err_files:
            src = r['file']
            stem, ext = os.path.splitext(src)
            dst = stem + '_符号修正' + ext
            try:
                n = F.write_docx(src, dst, F.active_fixes(False))
                print('  ✓ %s（%d 处）' % (dst, n))
            except OSError as exc:
                print('  ✗ 写失败：%s（%s）' % (src, exc))
        return 0

    if err_files:
        print('发现 %d 份 docx 含必修错字符。加 --fix 生成修正副本。' % len(err_files))
        for r in err_files[:10]:
            print('   %s' % r['file'])
        return 1
    print('未发现必修错字符。')
    return 0


if __name__ == '__main__':
    sys.exit(main())
