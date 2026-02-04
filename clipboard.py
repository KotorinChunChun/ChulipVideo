"""クリップボード操作モジュール.

Windows環境での画像クリップボードコピー機能を提供。
"""
from __future__ import annotations

import ctypes
import io
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from PIL import Image


def copy_image_to_clipboard(pil_img: "Image.Image") -> bool:
    """PIL画像をWindowsクリップボードにコピーする.
    
    Args:
        pil_img: コピーするPIL Image オブジェクト
        
    Returns:
        成功した場合True、失敗した場合False
    """
    # First try pywin32 if available (more reliable)
    try:
        import win32clipboard
        import win32con
        output = io.BytesIO()
        pil_img.convert('RGB').save(output, 'BMP')
        data = output.getvalue()[14:]
        output.close()
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_DIB, data)
        finally:
            win32clipboard.CloseClipboard()
        return True
    except Exception:
        pass

    # Fallback to ctypes on Windows to set CF_DIB data (BMP without BITMAPFILEHEADER)
    try:
        output = io.BytesIO()
        pil_img.convert('RGB').save(output, 'BMP')
        data = output.getvalue()[14:]
        output.close()

        GMEM_MOVEABLE = 0x0002
        CF_DIB = 8
        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32

        if not user32.OpenClipboard(0):
            raise RuntimeError('OpenClipboard failed')
        try:
            user32.EmptyClipboard()
            hGlobal = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
            if not hGlobal:
                raise RuntimeError('GlobalAlloc failed')
            pGlobal = kernel32.GlobalLock(hGlobal)
            if not pGlobal:
                kernel32.GlobalFree(hGlobal)
                raise RuntimeError('GlobalLock failed')
            # ensure destination pointer type
            ctypes.memmove(ctypes.c_void_p(pGlobal), data, len(data))
            kernel32.GlobalUnlock(hGlobal)
            user32.SetClipboardData(CF_DIB, hGlobal)
        finally:
            user32.CloseClipboard()
        return True
    except Exception as e:
        print(f"Clipboard copy failed: {e}")
        return False
