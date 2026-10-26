# -*- coding: utf-8 -*-
"""EMF/WMF -> PNG 转换器（跨平台）

EMF/WMF 是 Windows 图元格式，没有纯 Python 解码器，所以按平台分后端：

    Windows       系统 GDI+（ctypes 直调，零依赖）
    macOS/Linux   自动找 LibreOffice(soffice) 或 Inkscape 命令行
    都没有         不崩溃，给出人话提示 + 替代方案

命令行：
    python emf2png.py <源目录> <目标目录>

作为库：
    from emf2png import convert, backend_name
    convert('a.emf', 'a.png')      # 返回描述字符串，失败时含 'failed'

⚠️ 跨平台提示：老讲义的公式若是 MathType 的 WMF，**不必**转 PNG——
   `wmf_symbols.py` 能直接从 WMF 里读出文字符号，且完全跨平台。
"""
import os
import shutil
import subprocess
import sys

IS_WIN = sys.platform.startswith('win')

HINT_NO_BACKEND = """\
无法把 EMF/WMF 转成图片：当前系统（%s）没有可用的转换后端。

  · Windows 原生走 GDI+，无需额外安装；
  · macOS / Linux 需要下列任一个，装上后本脚本会自动调用：
        LibreOffice   macOS: brew install --cask libreoffice
                      Ubuntu: sudo apt install libreoffice
                      CentOS: sudo yum install libreoffice
        Inkscape      macOS: brew install --cask inkscape
                      Ubuntu: sudo apt install inkscape

  · 如果只是想要公式里的字符：不用转图，直接跑
        python wmf_symbols.py <文件或目录>
    它纯 Python 解析 WMF 矢量记录，Windows/macOS/Linux 都能用。
""" % sys.platform


# ---------------------------------------------------------------- GDI+ 后端
_gdiplus = None
_token = None


def _init_gdiplus():
    """仅 Windows：初始化 GDI+ 并绑定函数原型，返回模块对象"""
    global _gdiplus, _token
    if _gdiplus is not None:
        return _gdiplus
    if not IS_WIN:
        return None

    import ctypes
    from ctypes import wintypes, byref, c_void_p, c_uint, c_uint32, c_uint16, c_ubyte, c_int

    try:
        gdiplus = ctypes.WinDLL("gdiplus")
    except OSError:
        return None

    class GdiplusStartupInput(ctypes.Structure):
        _fields_ = [
            ("GdiplusVersion", c_uint32),
            ("DebugEventCallback", c_void_p),
            ("SuppressBackgroundThread", wintypes.BOOL),
            ("SuppressExternalCodecs", wintypes.BOOL),
        ]

    class GdiplusStartupOutput(ctypes.Structure):
        _fields_ = [("NotificationHook", c_void_p), ("NotificationUnhook", c_void_p)]

    class GUID(ctypes.Structure):
        _fields_ = [
            ("Data1", c_uint32),
            ("Data2", c_uint16),
            ("Data3", c_uint16),
            ("Data4", c_ubyte * 8),
        ]

    png_encoder = GUID(0x557CF406, 0x1A04, 0x11D3,
                       (c_ubyte * 8)(0x9A, 0x73, 0x00, 0x00, 0xF8, 0x1E, 0xF3, 0x2E))

    token = c_void_p()
    si = GdiplusStartupInput()
    si.GdiplusVersion = 1
    si.SuppressBackgroundThread = False
    si.SuppressExternalCodecs = False
    so = GdiplusStartupOutput()
    st = gdiplus.GdiplusStartup(byref(token), byref(si), byref(so))
    if st != 0:
        return None

    gdiplus.GdipLoadImageFromFile.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(c_void_p)]
    gdiplus.GdipLoadImageFromFile.restype = c_int
    gdiplus.GdipGetImageWidth.argtypes = [c_void_p, ctypes.POINTER(c_uint)]
    gdiplus.GdipGetImageHeight.argtypes = [c_void_p, ctypes.POINTER(c_uint)]
    gdiplus.GdipCreateBitmapFromScan0.argtypes = [
        c_int, c_int, c_int, c_int, c_void_p, ctypes.POINTER(c_void_p)
    ]
    gdiplus.GdipGetImageGraphicsContext.argtypes = [c_void_p, ctypes.POINTER(c_void_p)]
    gdiplus.GdipGraphicsClear.argtypes = [c_void_p, c_uint32]
    gdiplus.GdipDrawImageRectI.argtypes = [c_void_p, c_void_p, c_int, c_int, c_int, c_int]
    gdiplus.GdipSaveImageToFile.argtypes = [c_void_p, wintypes.LPCWSTR,
                                            ctypes.POINTER(GUID), c_void_p]
    gdiplus.GdipSetSmoothingMode.argtypes = [c_void_p, c_int]
    gdiplus.GdipDisposeImage.argtypes = [c_void_p]
    gdiplus.GdipDeleteGraphics.argtypes = [c_void_p]

    gdiplus._png_encoder = png_encoder
    gdiplus._token = token
    gdiplus._ctypes = ctypes
    _gdiplus = gdiplus
    _token = token
    return gdiplus


def _convert_windows(src, dst, scale, min_w, min_h, pad):
    g = _init_gdiplus()
    if g is None:
        return "failed: Windows 上 GDI+ 初始化失败（系统组件缺失？）"
    from ctypes import byref, c_void_p, c_uint
    try:
        return _convert_windows_impl(g, src, dst, scale, min_w, min_h, pad,
                                     byref, c_void_p, c_uint)
    except Exception as e:                                  # noqa: BLE001
        return "failed: %s" % e


def _convert_windows_impl(g, src, dst, scale, min_w, min_h, pad, byref, c_void_p, c_uint):
    img = c_void_p()
    r = g.GdipLoadImageFromFile(src, byref(img))
    if r != 0:
        return "failed: 读不了这个图元文件（GDI+ 码 %d）" % r

    w = c_uint()
    h = c_uint()
    g.GdipGetImageWidth(img, byref(w))
    g.GdipGetImageHeight(img, byref(h))

    ow = max(int(w.value * scale), min_w)
    oh = max(int(h.value * scale), min_h)
    tw, th = ow + pad * 2, oh + pad * 2

    bmp = c_void_p()
    r = g.GdipCreateBitmapFromScan0(tw, th, 0, 0x26200A, None, byref(bmp))
    if r != 0:
        g.GdipDisposeImage(img)
        return "failed: 建位图失败（GDI+ 码 %d）" % r

    gfx = c_void_p()
    g.GdipGetImageGraphicsContext(bmp, byref(gfx))
    g.GdipSetSmoothingMode(gfx, 4)
    g.GdipGraphicsClear(gfx, 0xFFFFFFFF)
    g.GdipDrawImageRectI(gfx, img, pad, pad, ow, oh)

    r = g.GdipSaveImageToFile(bmp, dst, byref(g._png_encoder), None)

    g.GdipDeleteGraphics(gfx)
    g.GdipDisposeImage(bmp)
    g.GdipDisposeImage(img)

    if r != 0:
        return "failed: 写 PNG 失败（GDI+ 码 %d）" % r
    return "%dx%d -> %dx%d" % (w.value, h.value, tw, th)


# ---------------------------------------------------- LibreOffice / Inkscape
def _find_bin(names):
    for n in names:
        p = shutil.which(n)
        if p:
            return p
    return None


_SOFFICE_CANDIDATES = ["soffice", "libreoffice",
                       "/Applications/LibreOffice.app/Contents/MacOS/soffice"]
_INKSCAPE_CANDIDATES = ["inkscape", "/Applications/Inkscape.app/Contents/MacOS/inkscape"]


def _convert_soffice(binpath, src, dst, scale):
    """soffice 只能按「目录 + 滤镜」批量转，这里转到临时目录再搬走"""
    import tempfile
    outdir = tempfile.mkdtemp(prefix='emf2png_')
    # PNG 导出分辨率靠 filter options 控制；不同版本参数名有差异，失败再退回默认
    tries = [
        ['--headless', '--norestore', '--convert-to',
         'png:impress_png_Export:{"PixelWidth":{"type":"long","value":%d}}' % (900 * scale // 3),
         '--outdir', outdir, src],
        ['--headless', '--norestore', '--convert-to', 'png', '--outdir', outdir, src],
    ]
    for args in tries:
        try:
            subprocess.run([binpath] + args, check=False, timeout=120,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:                                   # noqa: BLE001
            continue
        base = os.path.splitext(os.path.basename(src))[0] + '.png'
        cand = os.path.join(outdir, base)
        if os.path.exists(cand):
            shutil.move(cand, dst)
            shutil.rmtree(outdir, ignore_errors=True)
            return 'ok (LibreOffice)'
    shutil.rmtree(outdir, ignore_errors=True)
    return 'failed: LibreOffice 没能转出 PNG'


def _convert_inkscape(binpath, src, dst, scale):
    dpi = 96 * max(1, scale)
    try:
        p = subprocess.run([binpath, src, '--export-type=png',
                            '--export-filename=' + dst, '--export-dpi=%d' % dpi],
                           check=False, timeout=120,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:                                  # noqa: BLE001
        return 'failed: %s' % e
    if p.returncode == 0 and os.path.exists(dst):
        return 'ok (Inkscape)'
    return 'failed: Inkscape 返回码 %d' % p.returncode


def backend_name():
    """当前平台实际会用哪个后端，or None"""
    if _init_gdiplus() is not None:
        return 'GDI+ (Windows)'
    if _find_bin(_SOFFICE_CANDIDATES):
        return 'LibreOffice'
    if _find_bin(_INKSCAPE_CANDIDATES):
        return 'Inkscape'
    return None


def convert(src, dst, scale=3, min_w=240, min_h=90, pad=16):
    """EMF/WMF -> PNG。返回描述字符串；失败时以 'failed' 开头（不抛异常）"""
    if _init_gdiplus() is not None:
        return _convert_windows(src, dst, scale, min_w, min_h, pad)

    b = _find_bin(_SOFFICE_CANDIDATES)
    if b:
        return _convert_soffice(b, src, dst, scale)
    b = _find_bin(_INKSCAPE_CANDIDATES)
    if b:
        return _convert_inkscape(b, src, dst, scale)
    return 'failed: 没有可用的转换后端（详见 python emf2png.py --doctor）'


if __name__ == "__main__":
    if '--doctor' in sys.argv or len(sys.argv) < 3:
        print('平台      : %s' % sys.platform)
        print('转换后端  : %s' % (backend_name() or '无'))
        if backend_name() is None:
            print()
            print(HINT_NO_BACKEND)
        sys.exit(0)

    srcdir, dstdir = sys.argv[1], sys.argv[2]
    os.makedirs(dstdir, exist_ok=True)
    ok = 0
    total = 0
    for name in sorted(os.listdir(srcdir)):
        if not name.lower().endswith((".emf", ".wmf")):
            continue
        total += 1
        src = os.path.join(srcdir, name)
        dst = os.path.join(dstdir, os.path.splitext(name)[0] + ".png")
        res = convert(src, dst)
        print("%-24s %s" % (name, res))
        if not res.startswith('failed'):
            ok += 1
    print("转换成功 %d / %d" % (ok, total))
    if total and ok == 0 and backend_name() is None:
        print()
        print(HINT_NO_BACKEND)
        sys.exit(2)

    g = _init_gdiplus()
    if g is not None:
        try:
            g.GdiplusShutdown(g._token)
        except Exception:                                   # noqa: BLE001
            pass
