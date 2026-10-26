#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""doctor — latex-omml 环境体检

装完不确定好不好使？先跑这个：

    python scripts/doctor.py

它逐项检查 Python 版本、必需/可选依赖、内置转换表、平台能力，
最后跑一遍端到端自检。缺什么就直接给出对应的安装命令，
不用去翻文档猜。

    python scripts/doctor.py --json     # 机器可读，便于 CI
    python scripts/doctor.py --fix      # 直接 pip 装上缺的必需依赖
"""
import argparse
import importlib
import importlib.metadata as md
import json
import os
import platform
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))

# (import 名, pip 名, 是否必需, 用途)
DEPS = [
    ('lxml', 'lxml>=5.0', True, 'XML 变换（跑 MML2OMML.XSL）'),
    ('latex2mathml', 'latex2mathml>=3.80', True, 'LaTeX → MathML'),
    ('docx', 'python-docx>=1.1', True, '读写 docx'),
    ('pdfplumber', 'pdfplumber>=0.11', False, 'PDF 转 Word 用'),
    ('pymupdf', 'pymupdf', False, '扫描件 PDF 导出成图片用'),
    ('pypdfium2', 'pypdfium2', False, '同上，pymupdf 的备选'),
    ('olefile', 'olefile>=0.46', False, '解 MathType OLE 包（不装也能跑）'),
    ('PIL', 'pillow', False, '插图处理用'),
]

MIN_PY = (3, 9)

OK = 'OK'
MISS = '缺'
WARN = '注意'


def _ver(dist):
    try:
        return md.version(dist)
    except Exception:                                       # noqa: BLE001
        return '?'


def check_python():
    v = sys.version_info
    good = (v.major, v.minor) >= MIN_PY
    in_venv = (hasattr(sys, 'real_prefix')
               or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix))
    return {
        'item': 'Python',
        'value': '%d.%d.%d' % (v.major, v.minor, v.micro),
        'status': OK if good else MISS,
        'note': '要求 >= %d.%d；%s' % (*MIN_PY,
                                  '当前是虚拟环境' if in_venv else '当前是系统 Python'),
        'fix': '' if good else '装一个 3.9+ 的 Python',
    }


def check_deps():
    rows = []
    for mod, spec, required, purpose in DEPS:
        pip_name = spec.split('>')[0].split('=')[0].strip()
        try:
            m = importlib.import_module(mod)
            v = getattr(m, '__version__', None) or _ver(pip_name)
            rows.append({
                'item': pip_name, 'value': str(v),
                'status': OK,
                'note': ('必需 · ' if required else '可选 · ') + purpose,
                'fix': '',
            })
        except Exception as e:                              # noqa: BLE001
            rows.append({
                'item': pip_name, 'value': '—',
                'status': MISS if required else '可选缺',
                'note': ('必需 · ' if required else '可选 · ') + purpose,
                'fix': '' if required else '',
                'cause': type(e).__name__,
            })
    return rows


def check_xsl():
    """内置转换表：文件在不在，或者能不能从内嵌 bundle 还原出来"""
    p = os.path.join(HERE, 'MML2OMML.XSL')
    if os.path.exists(p):
        kb = os.path.getsize(p) / 1024
        return {'item': 'MML2OMML.XSL', 'value': '%.0f KB' % kb, 'status': OK,
                'note': '内置转换表（不需要装 Office）', 'fix': ''}
    try:
        sys.path.insert(0, HERE)
        import latex2omml as L
        got = L._extract_bundled_xsl()
        if got and os.path.exists(got):
            return {'item': 'MML2OMML.XSL', 'value': '已从内置包还原', 'status': OK,
                    'note': '内置转换表（不需要装 Office）', 'fix': ''}
    except Exception:                                       # noqa: BLE001
        pass
    return {'item': 'MML2OMML.XSL', 'value': '—', 'status': MISS,
            'note': '转换表缺失，公式无法生成',
            'fix': '重新安装本 skill（xsl_bundle.py 会在首次运行时自动还原）'}


def check_emf():
    try:
        sys.path.insert(0, HERE)
        import emf2png
        be = emf2png.backend_name()
    except Exception as e:                                  # noqa: BLE001
        return {'item': 'EMF/WMF 转图', 'value': '—', 'status': WARN,
                'note': '可选 · 老讲义的图元公式转图片', 'fix': str(e)}
    if be:
        return {'item': 'EMF/WMF 转图', 'value': be, 'status': OK,
                'note': '可选 · 老讲义的图元公式转图片', 'fix': ''}
    hint = ('Windows 走 GDI+；macOS/Linux 装 LibreOffice 或 Inkscape 后自动启用'
            if not sys.platform.startswith('win')
            else 'GDI+ 初始化失败')
    return {'item': 'EMF/WMF 转图', 'value': '不可用', 'status': '可选缺',
            'note': '可选 · 老讲义的图元公式转图片', 'fix': hint}


def check_end2end():
    try:
        sys.path.insert(0, HERE)
        import latex2omml as L
        ok, fail, detail = L.selftest()
        bad = [n for n, _, hit in detail if not hit]
        return {
            'item': '端到端自检', 'value': '%d / %d' % (ok, ok + fail),
            'status': OK if fail == 0 else MISS,
            'note': '核心符号全覆盖' + ('（失败：%s）' % '、'.join(bad) if bad else ''),
            'fix': '' if fail == 0 else '看 `python scripts/latex2omml.py` 的输出',
        }
    except Exception as e:                                  # noqa: BLE001
        return {'item': '端到端自检', 'value': '—', 'status': MISS,
                'note': '引擎跑不起来', 'fix': '%s: %s' % (type(e).__name__, e)}


def collect():
    return [check_python()] + check_deps() + [check_xsl(), check_emf(), check_end2end()]


def _dwidth(s):
    """终端显示宽度：中文/全角算 2 列，其余算 1 列"""
    import unicodedata
    w = 0
    for ch in s:
        w += 2 if unicodedata.east_asian_width(ch) in ('W', 'F') else 1
    return w


def _pad(s, width):
    return s + ' ' * max(0, width - _dwidth(s))


def print_report(rows):
    w = max(_dwidth(r['item']) for r in rows) + 2
    print('')
    print('latex-omml 环境体检    %s / Python %d.%d.%d'
          % (platform.system(), *sys.version_info[:3]))
    print('─' * 78)
    for r in rows:
        line = (_pad(r['item'], w) + _pad(r['value'], 18)
                + _pad(r['status'], 6) + r['note'])
        print(line)
        if r.get('fix'):
            print(_pad('', w + 2) + '→ %s' % r['fix'])
    print('─' * 78)

    required_missing = [r for r in rows
                        if r['status'] == MISS and '可选' not in r['note']]
    optional_missing = [r for r in rows if r['status'] == '可选缺']
    if required_missing:
        pkgs = [r['item'] for r in required_missing if r['item'] != 'MML2OMML.XSL']
        print('结论：还不能用，缺 %d 个必需组件。' % len(required_missing))
        if pkgs:
            print('      一行装齐：')
            print('        %s -m pip install %s'
                  % (os.path.basename(sys.executable), ' '.join(pkgs)))
        print('      或者直接跑：python scripts/doctor.py --fix')
    elif optional_missing:
        print('结论：核心功能可用。缺的是可选组件（%s），'
              '只影响 PDF 转 Word / 扫描件 / 图元这几个旁支。'
              % '、'.join(r['item'] for r in optional_missing))
    else:
        print('结论：全部就绪，可以直接用。')
    print('')


def fix(rows):
    pkgs = []
    for r in rows:
        if r['status'] == MISS and '可选' not in r['note'] and r['item'] != 'MML2OMML.XSL':
            spec = next((s for m, s, req, _ in DEPS if s.split('>')[0].split('=')[0] == r['item']),
                        r['item'])
            pkgs.append(spec)
    if not pkgs:
        print('没有需要安装的必需组件。')
        return 0
    cmd = [sys.executable, '-m', 'pip', 'install'] + pkgs
    print('执行：%s' % ' '.join(cmd))
    try:
        return subprocess.call(cmd)
    except Exception as e:                                  # noqa: BLE001
        print('安装失败：%s' % e)
        return 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--json', action='store_true')
    ap.add_argument('--fix', action='store_true')
    a = ap.parse_args()
    rows = collect()
    if a.json:
        print(json.dumps(rows, ensure_ascii=False, indent=2))
    else:
        print_report(rows)
    if a.fix:
        return fix(rows)
    bad = [r for r in rows if r['status'] == MISS and '可选' not in r['note']]
    return 1 if bad else 0


if __name__ == '__main__':
    sys.exit(main())
