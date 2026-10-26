#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把公式上标/下标/根指数的字号钳制到指定上限。

为什么需要它
------------
OMML 里 `m:sup` / `m:sub` / `m:deg` 内的 run 若显式写了 `w:sz`，
Word 就按**这个基准字号**再缩放渲染。于是同一份文档里：
12pt 段落（sz=24）的上标，永远比 11pt 段落（sz=22）的上标大一圈。
用户说「第三行上标比上面 x² 的 2 大」，要的就是把前者的基准压到和后者的参照物一致。

用法
----
    python clamp_script_sz.py 文件.docx                      # 全文，钳到 22
    python clamp_script_sz.py 文件.docx --paras 2            # 只改第 3 段（0 起）
    python clamp_script_sz.py 文件.docx --paras 0,2 --sz 20  # 指定段 + 指定上限
    python clamp_script_sz.py 文件.docx --inplace            # 原地改（自动 .bak）
    python clamp_script_sz.py 文件.docx --list               # 只看各段字号分布

安全
----
- 默认不覆盖原文件，输出 `原名_上标调整.docx`
- `--inplace` 先备份 `.bak.docx`
- 只动 sup/sub/deg **子树内**的 `w:sz`，底数、运算符、括号一律不碰
- 文件被 Word/WPS 占用时保留 `.new` 并提示，不会静默失败
"""
import argparse
import collections
import re
import shutil
import sys
import zipfile
from pathlib import Path

SCRIPT_TAGS = ('m:sup', 'm:sub', 'm:deg')
DEFAULT_SZ = 22  # 11pt：Word 里上标的常见参照值


def find_regions(s, open_t, close_t):
    """返回 [(start, end)]，end 为闭合标签之后的位置，支持嵌套。"""
    out = []
    pos = 0
    while True:
        a = s.find(open_t, pos)
        if a < 0:
            break
        depth = 0
        k = a
        end = None
        while k < len(s):
            b = s.find(open_t, k)
            c = s.find(close_t, k)
            if c < 0:
                break
            if b != -1 and b < c:
                depth += 1
                k = b + len(open_t)
            else:
                depth -= 1
                k = c + len(close_t)
                if depth == 0:
                    end = k
                    break
        if end is None:
            break
        out.append((a, end))
        pos = end
    return out


def para_regions(s):
    """定位所有顶层 w:p 段落区间（按出现顺序）。"""
    out = []
    pos = 0
    pat = re.compile(r'<w:p(?:\s[^>]*)?>|</w:p>')
    while True:
        m = pat.search(s, pos)
        if not m:
            break
        a = m.start()
        if m.group(0).startswith('</'):
            pos = m.end()
            continue
        tag_end = m.end()
        if s[a:tag_end].rstrip().endswith('/>'):
            out.append((a, tag_end))
            pos = tag_end
            continue
        depth = 1
        k = tag_end
        end = None
        while True:
            mm = pat.search(s, k)
            if not mm:
                break
            if mm.group(0).startswith('</'):
                depth -= 1
                k = mm.end()
                if depth == 0:
                    end = k
                    break
            else:
                k = mm.end()
                if not s[mm.start():mm.end()].rstrip().endswith('/>'):
                    depth += 1
        if end is None:
            break
        out.append((a, end))
        pos = end
    return out


def clamp_scripts(fragment, target=DEFAULT_SZ, exact=False):
    """把 fragment 内 sup/sub/deg 子树里的 w:sz 钳制到 <= target。

    exact=True 时改为「一律设为 target」（可升可降），便于在档位之间来回调。
    返回 (新片段, 改动次数)。
    """
    regions = []
    for tag in SCRIPT_TAGS:
        regions += find_regions(fragment, '<%s>' % tag, '</%s>' % tag)
    regions.sort()
    merged = []
    for a, b in regions:
        if merged and a <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], b))
        else:
            merged.append((a, b))

    total = 0
    for a, b in reversed(merged):  # 从后往前改，避免偏移
        seg = fragment[a:b]
        cnt = [0]

        def repl(m):
            val = int(m.group(2))
            hit = (val != target) if exact else (val > target)
            if hit:
                cnt[0] += 1
                return '%s%d%s' % (m.group(1), target, m.group(3))
            return m.group(0)

        seg2 = re.sub(r'(<w:sz w:val=")(\d+)("\s*/>)', repl, seg)
        if seg2 != seg:
            fragment = fragment[:a] + seg2 + fragment[b:]
            total += cnt[0]
    return fragment, total


def set_frac_sz(fragment, frac_sz):
    """把 sup/sub 子树里 m:f（分数）的 run 字号统一设为 frac_sz。

    为什么需要：Word 对 `m:sup` 里的**直接 run** 会再乘一次脚本缩放（约 0.65），
    但对 `m:f` 的 num/den **不缩放**。于是同一层里，`a^{m/n}` 的 m 会比
    `(√[n]{a})^m` 的 m 大一圈（实测 9px vs 6px，差 1.5 倍）。
    给分数单独补偿后两者才真正同号。

    经验值：frac_sz ≈ 目标 sz × 0.65。
    """
    sup_regions = []
    for tag in ('m:sup', 'm:sub'):
        sup_regions += find_regions(fragment, '<%s>' % tag, '</%s>' % tag)
    sup_regions.sort()
    merged = []
    for a, b in sup_regions:
        if merged and a <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], b))
        else:
            merged.append((a, b))

    total = 0
    for a, b in reversed(merged):
        seg = fragment[a:b]
        for fa, fb in reversed(find_regions(seg, '<m:f>', '</m:f>')):
            sub = seg[fa:fb]
            cnt = [0]

            def repl(m):
                if m.group(2) != str(frac_sz):
                    cnt[0] += 1
                    return '%s%d%s' % (m.group(1), frac_sz, m.group(3))
                return m.group(0)

            sub2 = re.sub(r'(<w:sz w:val=")(\d+)("\s*/>)', repl, sub)
            if sub2 != sub:
                seg = seg[:fa] + sub2 + seg[fb:]
                total += cnt[0]
        fragment = fragment[:a] + seg + fragment[b:]
    return fragment, total


def sz_stats(fragment):
    return dict(collections.Counter(re.findall(r'w:sz w:val="(\d+)"', fragment)))


def main():
    ap = argparse.ArgumentParser(description='钳制公式上标/根指数字号')
    ap.add_argument('path', help='目标 docx')
    ap.add_argument('--paras', default='all', help='段落索引，逗号分隔；all 表示全文')
    ap.add_argument('--sz', type=int, default=DEFAULT_SZ, help='目标字号上限（半磅），默认 22')
    ap.add_argument('--exact', action='store_true',
                    help='直接设为 --sz 的值（可升可降），默认只降不升')
    ap.add_argument('--frac-sz', type=int, default=0, dest='frac_sz',
                    help='上标/下标内分数的字号补偿（经验值：目标 sz × 0.65）；0 表示不处理')
    ap.add_argument('--inplace', action='store_true', help='原地修改（自动 .bak 备份）')
    ap.add_argument('--list', action='store_true', help='只列各段字号分布，不改')
    args = ap.parse_args()

    path = Path(args.path)
    if not path.is_file():
        raise SystemExit('文件不存在：%s' % path)

    z = zipfile.ZipFile(path)
    names = z.namelist()
    data = {n: z.read(n) for n in names}
    z.close()

    s = data['word/document.xml'].decode('utf-8')
    prs = para_regions(s)
    if not prs:
        raise SystemExit('没有找到段落，文件可能不是 docx')

    if args.list:
        print('%s：共 %d 段' % (path.name, len(prs)))
        for i, (a, b) in enumerate(prs):
            txt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', s[a:b]))[:40]
            print('  段%-3d sz=%-28s %r' % (i, sz_stats(s[a:b]), txt))
        return

    if args.paras.strip().lower() == 'all':
        targets = list(range(len(prs)))
    else:
        targets = [int(x) for x in args.paras.split(',') if x.strip() != '']
        bad = [t for t in targets if t >= len(prs)]
        if bad:
            raise SystemExit('段落索引越界：%s（共 %d 段）' % (bad, len(prs)))

    total = 0
    for i in sorted(targets, reverse=True):  # 从后往前改，保持偏移有效
        a, b = prs[i]
        frag2, n = clamp_scripts(s[a:b], args.sz, exact=args.exact)
        nf = 0
        if args.frac_sz:
            frag2, nf = set_frac_sz(frag2, args.frac_sz)
        if n or nf:
            s = s[:a] + frag2 + s[b:]
            total += n + nf
            msg = '  段 %d：%d 处上标/根指数 -> %d' % (i, n, args.sz)
            if nf:
                msg += '，%d 处分数 -> %d' % (nf, args.frac_sz)
            print(msg)
    if not total:
        print('没有需要改动的字号（已全部%s %d）'
              % ('等于' if args.exact else '不超过', args.sz))
        return

    out = path
    if not args.inplace:
        out = path.with_name(path.stem + '_上标调整.docx')
    else:
        bak = Path(str(path)[:-5] + '.bak.docx')
        try:
            shutil.copy2(path, bak)
            print('已备份：%s' % bak.name)
        except OSError as e:
            print('备份失败（不影响继续）：%s' % e)

    data['word/document.xml'] = s.encode('utf-8')
    tmp = Path(str(out) + '.new')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n in names:
            zo.writestr(n, data[n])
    try:
        tmp.replace(out)
    except PermissionError:
        print('\n[未能替换] %s 正被 Word/WPS 占用。' % out.name)
        print('  已保留：%s' % tmp.name)
        print('  关掉预览/Word 后执行： move /Y "%s" "%s"' % (tmp.name, out.name))
        sys.exit(2)
    print('已写入：%s（共 %d 处）' % (out, total))


if __name__ == '__main__':
    main()
