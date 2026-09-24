# -*- coding: utf-8 -*-
r"""
check_omml_delims — 定界符（m:d）体检 / 回归自测

背景（ECMA-376 + MML2OMML.XSL 的实际规则）
-----------------------------------------
m:d 是 OMML 的「定界符组」（括号、大括号、竖线…），它的
    m:begChr 默认值 = '('
    m:endChr 默认值 = ')'
XSL 只在**定界符不等于默认值**时才显式写出来。看清这三行就懂判据了：

    \left( ... \right)   -> 不写 begChr/endChr       渲染 '(' ')'   正确
    A=[-2,4)             -> 写 begChr='' + 不写 endChr 渲染 '' ')'    正确（半开区间）
    \begin{cases}        -> begChr='{' + endChr=''    渲染 '{' 无     正确

**踩过的坑**：曾经为了「让 Word 不画右括号」，把 cases 那个显式的
<m:endChr m:val=""/> 删掉了 —— 删掉等于「没声明」，Word 于是套用默认的 ')'，
分段函数右边就凭空多出一个撑大的右圆括号。**这个元素绝不能删。**

于是判据（也是本脚本的检测规则），分两级：
    必坏 ✗：begChr ∈ { | ‖ ⌈ ⌊ ⟨ … （天然右伴不是 ')' 的符号）且 endChr 缺失
            —— 这是本方流水线的特征伤（XSL 在 endChr≠默认值 ')' 时必显式写出，
               缺了只可能是被后处理删掉），实测就是分段函数右边多一个 ")"
    存疑 ?：其它非空 begChr（如 '['）且 endChr 缺失；或 endChr 非空但 begChr 缺失
            —— `[-2,4)` 半开区间是正常写法；WPS 产出的文档也大量缺 begChr，
               不一定坏，只提示复核
    正常：begChr='' 且 endChr 缺失（左显式不画、右用默认 ')'）
          两者都缺 / 两者都有

用法
----
    python check_omml_delims.py --selftest            # 回归自测（不需要 docx）
    python check_omml_delims.py <文件.docx>            # 扫成品，期望「必坏 0」
"""
import sys
import zipfile

from lxml import etree

M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'

# 天然右伴不是 ')' 的左定界符：配默认右括号必错
BAD_LEFT = set('{|‖⌈⌊⟨⌜⟦〖')


def scan_root(root):
    """返回 (必坏列表, 存疑列表)，元素为 (说明, 内容摘要)"""
    must, maybe = [], []
    for d in root.iter(M + 'd'):
        dPr = d.find(M + 'dPr')
        beg = dPr.find(M + 'begChr') if dPr is not None else None
        end = dPr.find(M + 'endChr') if dPr is not None else None
        bv = beg.get(M + 'val') if beg is not None else None
        ev = end.get(M + 'val') if end is not None else None
        txt = ''.join(t.text or '' for t in d.iter(M + 't'))[:60]

        if bv not in (None, '') and ev is None:
            why = f'左定界符 {bv!r} 但没声明右定界符 → Word 会补默认 ")"'
            (must if bv in BAD_LEFT else maybe).append((why, txt))
        elif ev not in (None, '') and bv is None:
            # 镜像情形：WPS 等其它工具产出的文档大量是这种写法（实测 2181 个 docx 里
            # 一票 WPS 文件都这样），不一定是坏的，只提示复核。
            why = f'右定界符 {ev!r} 但没声明左定界符 → 若右符不是 ")" 可能会补默认 "("'
            maybe.append((why, txt))
    return must, maybe


def scan_docx(path):
    root = etree.fromstring(zipfile.ZipFile(path).read('word/document.xml'))
    total = sum(1 for _ in root.iter(M + 'd'))
    must, maybe = scan_root(root)
    for why, txt in must:
        print(f'  ✗ {why}  | 内容: {txt!r}')
    for why, txt in maybe:
        print(f'  ? {why}（半开区间之类的正常写法，请自行确认）  | 内容: {txt!r}')
    print(f'定界符组 {total} 个，必坏 {len(must)} 个，存疑 {len(maybe)} 个')
    return len(must)


def selftest():
    """回归自测：
    1) cases 必须带显式空 endChr；
    2) 故意复现「删掉 endChr」的旧 bug，检测器必须报出来（负例测试）；
    3) 半开区间 [-2,4) 这种正常省略不得算必坏。"""
    import os
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from latex2omml import latex_to_omml

    ok = True

    cases = (
        r'\begin{cases}4k+1 & n=6k,k\in N^{*}\\ 4k+2 & n=6k+1,k\in N^{*}\\'
        r' 4k+3 & n=6k+2,k\in N^{*}\end{cases}'
    )
    xml = latex_to_omml(cases, 'block').decode('utf-8')

    if '<m:endChr m:val=""/>' not in xml:
        ok = False
        print('✗ cases 的 <m:endChr m:val=""/> 丢了 —— Word 会在右边补一个 ")"')

    root = etree.fromstring(xml.encode('utf-8'))
    if any(scan_root(root)):
        ok = False
        print('✗ 正常 cases 被误报')

    # 负例：手动复现旧 bug（删掉那个显式空 endChr），检测器必须抓到「必坏」
    buggy = etree.fromstring(xml.replace('<m:endChr m:val=""/>', '').encode('utf-8'))
    must, _ = scan_root(buggy)
    if not must:
        ok = False
        print('✗ 检测器漏报 —— 删掉 endChr 的 bug 没被抓成「必坏」')

    # 正常省略：半开区间不得算必坏
    half = etree.fromstring(latex_to_omml(r'A=[-2,4)', 'inline'))
    must_h, _ = scan_root(half)
    if must_h:
        ok = False
        print('✗ 半开区间 A=[-2,4) 被误判为必坏')

    print('回归自测:', 'PASS' if ok else 'FAIL')
    return 0 if ok else 1


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--selftest':
        sys.exit(selftest())
    if len(sys.argv) > 1:
        sys.exit(1 if scan_docx(sys.argv[1]) else 0)
    print(__doc__)
