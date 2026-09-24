from pathlib import Path
import numpy as np
from PIL import Image, ImageDraw

root = Path(r"C:\Users\32933\Documents\ChatGPT\PPT制图")
src = root / "超薄环路热管_炫酷科技版.png"
out = root / "超薄环路热管_PS分层素材"
out.mkdir(exist_ok=True)

im = Image.open(src).convert("RGBA")
rgb = np.asarray(im)[..., :3]
hsv = np.asarray(im.convert("HSV"))
h, s, v = hsv[..., 0], hsv[..., 1], hsv[..., 2]
H, W = h.shape
yy, xx = np.mgrid[0:H, 0:W]

# Color layers constrained to the physical device region, avoiding decorative traces.
device_zone = (yy > 285) | ((xx > 1080) & (yy > 240))
blue = (h >= 82) & (h <= 148) & (s > 75) & (v > 45) & device_zone
red = ((h < 8) | (h > 247)) & (s > 95) & (v > 38) & (xx > 1080) & (yy > 230)
orange = (h >= 5) & (h <= 40) & (s > 80) & (v > 48) & device_zone & ~blue & ~red

def polygon_mask(points):
    m = Image.new("L", (W, H), 0)
    ImageDraw.Draw(m).polygon(points, fill=255)
    return np.asarray(m) > 0

# Broad component zones; color layers take priority so the composite remains clean.
display = polygon_mask([(0,120),(760,65),(1055,105),(1055,500),(945,555),(0,595)])
board = polygon_mask([(500,55),(1325,0),(1671,75),(1671,735),(1180,940),(610,780)])
occupied = blue | red | orange
display &= ~occupied
board &= ~occupied & ~display
background = ~(occupied | display | board)

layers = [
    ("01_环境背景.png", background),
    ("02_主板与电子元件.png", board),
    ("03_黑色显示承载区.png", display),
    ("04_橙铜超薄环路热管.png", orange),
    ("05_蓝色工质区域.png", blue),
    ("06_芯片红色热区.png", red),
]

base = np.asarray(im).copy()
for filename, mask in layers:
    arr = base.copy()
    arr[..., 3] = np.where(mask, base[..., 3], 0).astype(np.uint8)
    Image.fromarray(arr, "RGBA").save(out / filename)

# Alpha masks are useful for selections, adjustment layers, and manual refinement.
for filename, mask in layers:
    Image.fromarray((mask * 255).astype(np.uint8), "L").save(out / ("蒙版_" + filename))

print(f"{W}x{H}")
print(out)
