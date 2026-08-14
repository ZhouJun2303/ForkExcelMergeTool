# One-off helper: build ExcelMergeFork.ico from the generated master tile.
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parents[3]
MASTER = Path(
    r"C:\Users\37078\.grok\sessions\D%3A%5CMyGit%5CForkExcelMergeTool\019fff11-1f67-7ae3-9f83-c268595338f1\images\1.jpg"
)
OUT_ICO = ROOT / "Assets" / "ExcelMergeFork.ico"
OUT_16 = ROOT / "Assets" / "icons" / "app_16.png"
OUT_32 = ROOT / "Assets" / "icons" / "app_32.png"
OUT_256 = ROOT / "Assets" / "icons" / "app_256.png"


def punch_corners(img: Image.Image) -> Image.Image:
    rgba = img.convert("RGBA")
    px = rgba.load()
    w, h = rgba.size

    def is_halo(x: int, y: int) -> bool:
        r, g, b, a = px[x, y]
        return a > 0 and r > 228 and g > 228 and b > 228

    stack = [(0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1)]
    seen = set()
    while stack:
        x, y = stack.pop()
        if (x, y) in seen or x < 0 or y < 0 or x >= w or y >= h:
            continue
        seen.add((x, y))
        if not is_halo(x, y):
            continue
        px[x, y] = (0, 0, 0, 0)
        stack.extend(((x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)))
    return rgba


def fit_square(img: Image.Image, size: int) -> Image.Image:
    canvas = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    scaled = img.copy()
    scaled.thumbnail((size, size), Image.Resampling.LANCZOS)
    x = (size - scaled.width) // 2
    y = (size - scaled.height) // 2
    canvas.paste(scaled, (x, y), scaled)
    return canvas


def main() -> None:
    src = punch_corners(Image.open(MASTER))
    master = fit_square(src, 256)
    OUT_ICO.parent.mkdir(parents=True, exist_ok=True)
    master.save(
        OUT_ICO,
        format="ICO",
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )
    fit_square(src, 16).save(OUT_16)
    fit_square(src, 32).save(OUT_32)
    master.save(OUT_256)
    print("wrote", OUT_ICO, OUT_16, OUT_32, OUT_256)


if __name__ == "__main__":
    main()
