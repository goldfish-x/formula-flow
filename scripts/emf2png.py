"""EMF/WMF -> PNG converter using Windows GDI+ via ctypes (no .NET, no extra deps)."""
import ctypes
import os
import sys
from ctypes import wintypes, byref, c_void_p, c_uint, c_uint32, c_uint16, c_ubyte, c_int

gdiplus = ctypes.WinDLL("gdiplus")


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


PNG_ENCODER = GUID(0x557CF406, 0x1A04, 0x11D3, (c_ubyte * 8)(0x9A, 0x73, 0x00, 0x00, 0xF8, 0x1E, 0xF3, 0x2E))

token = c_void_p()
si = GdiplusStartupInput()
si.GdiplusVersion = 1
si.SuppressBackgroundThread = False
si.SuppressExternalCodecs = False
so = GdiplusStartupOutput()
st = gdiplus.GdiplusStartup(byref(token), byref(si), byref(so))
if st != 0:
    print("GdiplusStartup failed:", st)
    sys.exit(1)

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
gdiplus.GdipSaveImageToFile.argtypes = [c_void_p, wintypes.LPCWSTR, ctypes.POINTER(GUID), c_void_p]
gdiplus.GdipSetSmoothingMode.argtypes = [c_void_p, c_int]
gdiplus.GdipDisposeImage.argtypes = [c_void_p]
gdiplus.GdipDeleteGraphics.argtypes = [c_void_p]


def convert(src, dst, scale=3, min_w=240, min_h=90, pad=16):
    img = c_void_p()
    r = gdiplus.GdipLoadImageFromFile(src, byref(img))
    if r != 0:
        return "load failed (%d)" % r

    w = c_uint()
    h = c_uint()
    gdiplus.GdipGetImageWidth(img, byref(w))
    gdiplus.GdipGetImageHeight(img, byref(h))

    ow = max(int(w.value * scale), min_w)
    oh = max(int(h.value * scale), min_h)
    tw, th = ow + pad * 2, oh + pad * 2

    bmp = c_void_p()
    r = gdiplus.GdipCreateBitmapFromScan0(tw, th, 0, 0x26200A, None, byref(bmp))
    if r != 0:
        gdiplus.GdipDisposeImage(img)
        return "bitmap failed (%d)" % r

    gfx = c_void_p()
    gdiplus.GdipGetImageGraphicsContext(bmp, byref(gfx))
    gdiplus.GdipSetSmoothingMode(gfx, 4)
    gdiplus.GdipGraphicsClear(gfx, 0xFFFFFFFF)
    gdiplus.GdipDrawImageRectI(gfx, img, pad, pad, ow, oh)

    r = gdiplus.GdipSaveImageToFile(bmp, dst, byref(PNG_ENCODER), None)

    gdiplus.GdipDeleteGraphics(gfx)
    gdiplus.GdipDisposeImage(bmp)
    gdiplus.GdipDisposeImage(img)

    if r != 0:
        return "save failed (%d)" % r
    return "%dx%d -> %dx%d" % (w.value, h.value, tw, th)


if __name__ == "__main__":
    srcdir, dstdir = sys.argv[1], sys.argv[2]
    os.makedirs(dstdir, exist_ok=True)
    ok = 0
    for name in sorted(os.listdir(srcdir)):
        if not name.lower().endswith((".emf", ".wmf")):
            continue
        src = os.path.join(srcdir, name)
        dst = os.path.join(dstdir, os.path.splitext(name)[0] + ".png")
        res = convert(src, dst)
        print("%-16s %s" % (name, res))
        if "failed" not in res:
            ok += 1
    print("converted %d files" % ok)
    gdiplus.GdiplusShutdown(token)
