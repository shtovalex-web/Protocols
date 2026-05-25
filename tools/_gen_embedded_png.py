# -*- coding: utf-8 -*-
"""Одноразовая генерация ProtocolOHT_next/protocol_embedded_assets.py"""
from __future__ import annotations

import base64
import math
import struct
import zlib
from pathlib import Path


def rgba_png(w: int, h: int, pixels: bytes) -> bytes:
    sig = b"\x89PNG\r\n\x1a\n"

    def chunk(t: bytes, d: bytes) -> bytes:
        crc = zlib.crc32(t + d) & 0xFFFFFFFF
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", crc)

    ihdr = struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0)
    raw = b""
    for y in range(h):
        raw += b"\x00" + pixels[y * w * 4 : (y + 1) * w * 4]
    return sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b"")


def main() -> None:
    w = 64
    px = bytearray(w * w * 4)
    for y in range(w):
        for x in range(w):
            i = (y * w + x) * 4
            px[i : i + 4] = bytes([0x1A, 0x5F, 0xB4, 255])
    for y in range(12, 53):
        for x in range(14, 23):
            i = (y * w + x) * 4
            px[i : i + 4] = bytes([255, 255, 255, 255])
    for y in (18, 24, 30):
        for x in range(26, 49):
            i = (y * w + x) * 4
            px[i : i + 4] = bytes([0x1A, 0x5F, 0xB4, 255])
    logo_b64 = base64.b64encode(rgba_png(w, w, bytes(px))).decode("ascii")

    w2 = 22
    px2 = bytearray(w2 * w2 * 4)
    cx, cy = 10.5, 10.5
    for y in range(w2):
        for x in range(w2):
            i = (y * w2 + x) * 4
            d = math.hypot(x - cx, y - cy)
            if d <= 10:
                px2[i : i + 4] = bytes([240, 244, 248, 255])
            else:
                px2[i : i + 4] = bytes([0, 0, 0, 0])
    for ang in range(0, 270, 6):
        r = 5.5
        xi = int(11 + r * math.cos(math.radians(ang)))
        yi = int(11 + r * math.sin(math.radians(ang)))
        for dx, dy in ((0, 0), (1, 0), (-1, 0), (0, 1), (0, -1)):
            xa, ya = xi + dx, yi + dy
            if 0 <= xa < w2 and 0 <= ya < w2:
                j = (ya * w2 + xa) * 4
                if px2[j + 3]:
                    px2[j : j + 4] = bytes([0x22, 0x55, 0x88, 255])
    for x, y in ((16, 6), (17, 5), (18, 4), (17, 7), (18, 6)):
        if 0 <= x < w2 and 0 <= y < w2:
            j = (y * w2 + x) * 4
            if px2[j + 3]:
                px2[j : j + 4] = bytes([0x22, 0x55, 0x88, 255])
    ref_b64 = base64.b64encode(rgba_png(w2, w2, bytes(px2))).decode("ascii")

    out = Path(__file__).resolve().parent.parent / "ProtocolOHT_next" / "protocol_embedded_assets.py"
    body = f'''# -*- coding: utf-8 -*-
"""Встроенные изображения интерфейса (логотип, кнопка обновления баз). Перегенерация: py -3 tools/_gen_embedded_png.py"""

from __future__ import annotations

import base64

_EMBEDDED_LOGO_PNG_B64 = "{logo_b64}"

_EMBEDDED_REFRESH_BTN_PNG_B64 = "{ref_b64}"


def embedded_logo_png_bytes() -> bytes:
    return base64.b64decode(_EMBEDDED_LOGO_PNG_B64)


def embedded_refresh_button_png_bytes() -> bytes:
    return base64.b64decode(_EMBEDDED_REFRESH_BTN_PNG_B64)
'''
    out.write_text(body, encoding="utf-8")
    print("written", out, "logo", len(logo_b64), "ref", len(ref_b64))


if __name__ == "__main__":
    main()
