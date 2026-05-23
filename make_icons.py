"""One-off icon generator for the PWA. Run once after install."""
from PIL import Image, ImageDraw, ImageFont

BG = (15, 15, 16)        # dark background matching theme
FG = (200, 156, 90)      # sepia accent matching theme


def make_icon(size: int, out_path: str) -> None:
    img = Image.new("RGB", (size, size), BG)
    draw = ImageDraw.Draw(img)

    # Rounded rect background
    pad = int(size * 0.08)
    draw.rounded_rectangle(
        [pad, pad, size - pad, size - pad],
        radius=int(size * 0.18),
        fill=BG,
        outline=FG,
        width=max(3, size // 80),
    )

    # AEY text
    text = "AEY"
    try:
        font = ImageFont.truetype("arialbd.ttf", int(size * 0.32))
    except OSError:
        font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    x = (size - tw) // 2 - bbox[0]
    y = (size - th) // 2 - bbox[1] - int(size * 0.02)
    draw.text((x, y), text, fill=FG, font=font)

    img.save(out_path, "PNG")
    print(f"Saved {out_path}")


if __name__ == "__main__":
    make_icon(192, r"C:\Users\CREATOR TECH\Desktop\aey-web\static\icon-192.png")
    make_icon(512, r"C:\Users\CREATOR TECH\Desktop\aey-web\static\icon-512.png")
