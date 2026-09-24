from PIL import Image, ImageDraw
from pathlib import Path

root = Path(r'C:\Users\32933\Documents\ChatGPT\PPT制图\build')
im = Image.open(root / 'source123_landscape.png').convert('RGB')
# Remove only the five movable graphics/text groups; all red production geometry stays unchanged.
logo = im.crop((1022, 285, 1115, 375))
draw = ImageDraw.Draw(im)
for box in [
    (1009, 271, 1128, 385),  # university seal, not the Innovation Harbour mark
    (325, 663, 1222, 770),   # device title and its former frame
    (556, 803, 922, 989),   # project diamond
    (466, 1025, 965, 1147), # leader and member credits
]:
    draw.rectangle(box, fill='white')
im.save(root / 'rearrange_123' / 'fixed_panel.png')
logo.save(root / 'rearrange_123' / 'school_seal.png')
