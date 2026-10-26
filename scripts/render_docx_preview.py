#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx -> PDF -> PNG，用于公式排版的视觉核对（Word COM 版）。

用途
----
改完公式字号/定界符后，把它渲染成 PNG，肉眼确认或做像素测量，
不必让用户自己开 Word 一张张看。

用法
----
    python render_docx_preview.py 输出目录 文件1.docx [文件2.docx ...]
    python render_docx_preview.py 输出目录 文件.docx --dpi 600

依赖
----
    pip install pywin32 pymupdf

⚠️ 已知限制（2026-10-26 实测）
----------------------------
某些 Windows 环境（含本机 Office16）`Word.Application` 的 COM 无法启动：
    pywintypes.com_error: (-2146959355, '服务器运行失败', None, None)
    = CO_E_SERVER_EXEC_FAILURE
即使把命令放到沙箱外执行、或用 DispatchEx / Dispatch 都一样，属于会话隔离限制，
不是代码问题。此时脚本会明确报出来，请改走：
    ① 量用户提供的截图（墨迹高度比）；
    ② 直接请用户在 Word 里瞄一眼给结论。
若机器上有 LibreOffice，也可用 `soffice --headless --convert-to pdf` 替代本脚本的
第一步（emf2png.py 里有现成的后端探测代码可复用）。
"""
import argparse
import sys
from pathlib import Path

WD_EXPORT_PDF = 17


def to_pdf(docx: Path, pdf: Path) -> Path:
    import win32com.client as win32

    try:
        word = win32.DispatchEx('Word.Application')
    except Exception as e:  # com_error
        raise RuntimeError(
            '无法启动 Word COM（CO_E_SERVER_EXEC_FAILURE 表示当前会话被隔离）。\n'
            '原始错误：%s' % e
        )
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(docx), ReadOnly=True, AddToRecentFiles=False)
        try:
            doc.ExportAsFixedFormat(
                OutputFileName=str(pdf),
                ExportFormat=WD_EXPORT_PDF,
                OpenAfterExport=False,
                OptimizeFor=0,
                Range=0,
                Item=0,
                IncludeDocProps=False,
                KeepIRM=False,
                CreateBookmarks=0,
                DocStructureTags=False,
                BitmapMissingFonts=True,
                UseISO19005_1=False,
            )
        finally:
            doc.Close(SaveChanges=0)
    finally:
        word.Quit()
    return pdf


def pdf_to_png(pdf: Path, outdir: Path, dpi=400):
    import pymupdf

    d = pymupdf.open(str(pdf))
    out = []
    for i, page in enumerate(d):
        pix = page.get_pixmap(dpi=dpi)
        p = outdir / ('%s_p%d.png' % (pdf.stem, i + 1))
        pix.save(str(p))
        out.append((p, pix.width, pix.height))
    d.close()
    return out


def main():
    ap = argparse.ArgumentParser(description='docx 渲染成 PNG 做视觉核对')
    ap.add_argument('outdir')
    ap.add_argument('files', nargs='+')
    ap.add_argument('--dpi', type=int, default=400)
    ap.add_argument('--keep-pdf', action='store_true')
    args = ap.parse_args()

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    failed = 0
    for arg in args.files:
        docx = Path(arg)
        pdf = outdir / (docx.stem + '.pdf')
        try:
            to_pdf(docx, pdf)
            for p, w, h in pdf_to_png(pdf, outdir, args.dpi):
                print('OK   %s  %dx%d' % (p, w, h))
            if not args.keep_pdf:
                pdf.unlink(missing_ok=True)
        except Exception as e:
            failed += 1
            print('FAIL %s\n     %s' % (docx.name, e), file=sys.stderr)
    sys.exit(1 if failed else 0)


if __name__ == '__main__':
    main()
