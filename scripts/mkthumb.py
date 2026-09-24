# -*- coding: utf-8 -*-
"""批量生成缩略图（GDI+ 直接调用，绕开 PowerShell 策略限制）"""
import ctypes, os, sys
from ctypes import wintypes, byref, c_void_p, c_uint32, c_uint16, c_ubyte, c_int

gdiplus = ctypes.WinDLL('gdiplus')

class GdiplusStartupInput(ctypes.Structure):
    _fields_ = [('GdiplusVersion', c_uint32), ('DebugEventCallback', c_void_p),
                ('SuppressBackgroundThread', wintypes.BOOL), ('SuppressExternalCodecs', wintypes.BOOL)]

class GUID(ctypes.Structure):
    _fields_ = [('Data1', c_uint32), ('Data2', c_uint16), ('Data3', c_uint16), ('Data4', c_ubyte * 8)]

PNG = GUID(0x557CF406, 0x1A04, 0x11D3, (c_ubyte * 8)(0x9A, 0x73, 0x00, 0x00, 0xF8, 0x1E, 0xF3, 0x2E))
si = GdiplusStartupInput(); si.GdiplusVersion = 1
token = c_void_p()
gdiplus.GdiplusStartup(byref(token), byref(si), None)

gdiplus.GdipLoadImageFromFile.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(c_void_p)]
gdiplus.GdipLoadImageFromFile.restype = c_int
gdiplus.GdipGetImageWidth.argtypes = [c_void_p, ctypes.POINTER(c_uint32)]
gdiplus.GdipGetImageHeight.argtypes = [c_void_p, ctypes.POINTER(c_uint32)]
gdiplus.GdipCreateBitmapFromScan0.argtypes = [c_int, c_int, c_int, c_int, c_void_p, ctypes.POINTER(c_void_p)]
gdiplus.GdipGetImageGraphicsContext.argtypes = [c_void_p, ctypes.POINTER(c_void_p)]
gdiplus.GdipGraphicsClear.argtypes = [c_void_p, c_uint32]
gdiplus.GdipSetInterpolationMode = getattr(gdiplus, 'GdipSetInterpolationMode', None)
gdiplus.GdipDrawImageRectI.argtypes = [c_void_p, c_void_p, c_int, c_int, c_int, c_int]
gdiplus.GdipSaveImageToFile.argtypes = [c_void_p, wintypes.LPCWSTR, ctypes.POINTER(GUID), c_void_p]
gdiplus.GdipDisposeImage.argtypes = [c_void_p]
gdiplus.GdipDeleteGraphics.argtypes = [c_void_p]


def thumb(src, dst, max_w=1000, max_h=420, pad=16):
    img = c_void_p()
    if gdiplus.GdipLoadImageFromFile(src, byref(img)) != 0:
        return None
    w = c_uint32(); h = c_uint32()
    gdiplus.GdipGetImageWidth(img, byref(w))
    gdiplus.GdipGetImageHeight(img, byref(h))
    scale = min((max_w - pad * 2) / max(w.value, 1), (max_h - pad * 2) / max(h.value, 1), 3.0)
    tw, th = max(int(w.value * scale), 8), max(int(h.value * scale), 8)
    cw, ch = tw + pad * 2, th + pad * 2
    bmp = c_void_p()
    gdiplus.GdipCreateBitmapFromScan0(cw, ch, 0, 0x26200A, None, byref(bmp))
    gfx = c_void_p()
    gdiplus.GdipGetImageGraphicsContext(bmp, byref(gfx))
    gdiplus.GdipGraphicsClear(gfx, 0xFFFFFFFF)
    gdiplus.GdipDrawImageRectI(gfx, img, pad, pad, tw, th)
    gdiplus.GdipSaveImageToFile(bmp, dst, byref(PNG), None)
    gdiplus.GdipDeleteGraphics(gfx)
    gdiplus.GdipDisposeImage(bmp)
    gdiplus.GdipDisposeImage(img)
    return (w.value, h.value, cw, ch)


if __name__ == '__main__':
    src_dir, dst_dir = sys.argv[1], sys.argv[2]
    os.makedirs(dst_dir, exist_ok=True)
    rows = []
    for f in sorted(os.listdir(src_dir)):
        if not f.lower().endswith('.png'):
            continue
        r = thumb(os.path.join(src_dir, f), os.path.join(dst_dir, f))
        if r:
            w, h, cw, ch = r
            rows.append((f, w, h, round(w / max(h, 1), 2)))
    rows.sort(key=lambda t: int(''.join(c for c in t[0] if c.isdigit()) or 0))
    print(f'{"文件":<22}{"原始":>14}{"宽高比":>9}')
    for f, w, h, ar in rows:
        flag = '  <== 扁(可能跨行)' if ar > 2.6 else ''
        print(f'{f:<22}{w:>7}x{h:<6}{ar:>9}{flag}')
    gdiplus.GdiplusShutdown(token)
