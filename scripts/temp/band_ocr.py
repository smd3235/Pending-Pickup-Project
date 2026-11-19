import os
import subprocess

from PIL import Image

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
temp = os.path.join(root, "scripts", "temp")
src = os.path.join(temp, "shots2", "02_wndclass_desked_gsk.png")

img = Image.open(src).convert("RGB")
w, h = img.size
px = img.load()

targets = [(0, 0, 128), (0, 48, 144), (0, 112, 208), (16, 48, 144), (0, 96, 176)]
rows = []
for y in range(h):
    cnt = 0
    for x in range(w // 3, w):
        r, g, b = px[x, y]
        for tr, tg, tb in targets:
            if abs(r - tr) <= 24 and abs(g - tg) <= 24 and abs(b - tb) <= 24:
                cnt += 1
                break
    if cnt > 60:
        rows.append(y)

bands = []
for y in rows:
    if bands and y - bands[-1][1] <= 3:
        bands[-1][1] = y
    else:
        bands.append([y, y])
print("bands:", [(a, b, b - a) for a, b in bands])

out = subprocess.run(
    ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", os.path.join(temp, "ocr_boxes.ps1"), "-Path", src],
    capture_output=True, text=True,
)
words = []
for line in out.stdout.splitlines():
    if "\t" not in line:
        continue
    box, text = line.split("\t", 1)
    x, y, ww, hh = (int(v) for v in box.split(","))
    words.append((x, y, ww, hh, text))

print("ocr words:", len(words))
for band in bands:
    y0, y1 = band
    picked = [wd for wd in words if y0 - 4 <= wd[1] + wd[3] / 2 <= y1 + 4]
    picked.sort(key=lambda wd: wd[0])
    print(f"--- band {y0}-{y1}: {' '.join(p[4] for p in picked)}")
