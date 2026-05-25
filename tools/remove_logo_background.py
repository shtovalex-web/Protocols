# -*- coding: utf-8 -*-
"""
Убирает однотонный светлый фон у bundle/program_logo.png (заливка от краёв),
сохраняет PNG с альфой и обновляет вшитый base64 в protocol_embedded_assets.py.

После замены файла bundle/program_logo.png запускайте из корня проекта:
  py -3 tools/remove_logo_background.py
"""
from __future__ import annotations

import base64
from collections import deque
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
BUNDLE_LOGO = ROOT / "bundle" / "program_logo.png"
ASSETS_PY = ROOT / "ProtocolOHT_next" / "protocol_embedded_assets.py"


def remove_edge_connected_light_background(
    img: Image.Image,
    *,
    rgb_min: int = 248,
) -> Image.Image:
    """Прозрачность для пикселей, связанных с краем и достаточно светлых (фон)."""
    rgba = img.convert("RGBA")
    w, h = rgba.size
    px = rgba.load()

    def is_bg(r: int, g: int, b: int, a: int) -> bool:
        if a < 10:
            return True
        return r >= rgb_min and g >= rgb_min and b >= rgb_min

    visited = [[False] * w for _ in range(h)]
    q: deque[tuple[int, int]] = deque()

    for x in range(w):
        if is_bg(*px[x, 0]):
            q.append((x, 0))
        if is_bg(*px[x, h - 1]):
            q.append((x, h - 1))
    for y in range(h):
        if is_bg(*px[0, y]):
            q.append((0, y))
        if is_bg(*px[w - 1, y]):
            q.append((w - 1, y))

    while q:
        x, y = q.popleft()
        if visited[y][x]:
            continue
        r, g, b, a = px[x, y]
        if not is_bg(r, g, b, a):
            continue
        visited[y][x] = True
        px[x, y] = (255, 255, 255, 0)
        for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
            nx, ny = x + dx, y + dy
            if 0 <= nx < w and 0 <= ny < h and not visited[ny][nx]:
                q.append((nx, ny))

    return rgba


def patch_embedded_assets(png_bytes: bytes) -> None:
    b64 = base64.b64encode(png_bytes).decode("ascii")
    text = ASSETS_PY.read_text(encoding="utf-8")
    marker = '_EMBEDDED_LOGO_PNG_B64 = "'
    i = text.find(marker)
    if i < 0:
        raise SystemExit("Не найден _EMBEDDED_LOGO_PNG_B64 в protocol_embedded_assets.py")
    j = text.find('"', i + len(marker))
    if j < 0:
        raise SystemExit("Не закрыта кавычка у _EMBEDDED_LOGO_PNG_B64")
    new_text = text[: i + len(marker)] + b64 + text[j:]
    ASSETS_PY.write_text(new_text, encoding="utf-8")


def main() -> int:
    if not BUNDLE_LOGO.is_file():
        print("Нет файла:", BUNDLE_LOGO)
        return 1
    img = Image.open(BUNDLE_LOGO)
    out = remove_edge_connected_light_background(img, rgb_min=248)
    out.save(BUNDLE_LOGO, format="PNG", optimize=True)
    png_bytes = BUNDLE_LOGO.read_bytes()
    print("Записано:", BUNDLE_LOGO, "размер", len(png_bytes), "байт")
    patch_embedded_assets(png_bytes)
    print("Обновлён:", ASSETS_PY)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
