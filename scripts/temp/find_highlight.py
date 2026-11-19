import os
from collections import Counter

from PIL import Image

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
src = os.path.join(root, "scripts", "temp", "shots2", "02_wndclass_desked_gsk.png")
img = Image.open(src).convert("RGB")
w, h = img.size
print("size", w, h)

px = img.load()
counts = Counter()
for y in range(0, h, 2):
    for x in range(0, w, 2):
        r, g, b = px[x, y]
        if b > 120 and b - r > 40 and b - g > 30:
            counts[(r // 16 * 16, g // 16 * 16, b // 16 * 16)] += 1
print("top blue-ish colors:", counts.most_common(8))
