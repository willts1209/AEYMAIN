"""One-shot helper: render Flushing_Farm_activity_log PDF to a single tall PNG.

Stitches both pages vertically, crops trailing blank space from the second
page so the result is a single continuous-looking table image suitable for
the landing-page hero figure.
"""
import pypdfium2 as pdfium
from PIL import Image

PDF_PATH = r"C:\Users\CREATOR TECH\Downloads\Flushing_Farm_activity_log.docx.pdf"
OUT_PATH = r"C:\Users\CREATOR TECH\Desktop\aey-web\static\img\fieldcast-output.png"
SCALE = 3  # ~216 DPI, sharp on retina displays

pdf = pdfium.PdfDocument(PDF_PATH)
imgs = [page.render(scale=SCALE).to_pil().convert("RGB") for page in pdf]

# Find first and last non-white rows of each page, return cropped to content
# plus a small margin so adjacent pages butt up cleanly when stitched.
def crop_to_content(img: Image.Image, top_margin: int = 24, bottom_margin: int = 40) -> Image.Image:
    px = img.load()
    w, h = img.size
    sample_xs = range(0, w, max(1, w // 60))

    def row_is_white(y: int) -> bool:
        for x in sample_xs:
            r, g, b = px[x, y]
            if r < 248 or g < 248 or b < 248:
                return False
        return True

    first = 0
    for y in range(h):
        if not row_is_white(y):
            first = y
            break
    last = h - 1
    for y in range(h - 1, -1, -1):
        if not row_is_white(y):
            last = y
            break
    top = max(0, first - top_margin)
    bottom = min(h, last + bottom_margin)
    return img.crop((0, top, w, bottom))

imgs = [crop_to_content(im) for im in imgs]

# Stitch vertically using the wider page as the canvas width, with a small
# gap between pages so the page break is visually subtle.
GAP = 12
w = max(im.width for im in imgs)
h = sum(im.height for im in imgs) + GAP * (len(imgs) - 1)
canvas = Image.new("RGB", (w, h), "white")
y = 0
for i, im in enumerate(imgs):
    x = (w - im.width) // 2
    canvas.paste(im, (x, y))
    y += im.height + (GAP if i < len(imgs) - 1 else 0)

canvas.save(OUT_PATH, "PNG", optimize=True)
print(f"Wrote {OUT_PATH}: {canvas.size[0]}x{canvas.size[1]} px")
