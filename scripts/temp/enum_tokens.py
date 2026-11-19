import os
import re

d = r"C:\Users\smd32\Downloads\Pending Pickup Project\vba"
pat = re.compile(r"(?<![\w.])((?:xl|mso|vb|pp|office)[A-Za-z0-9_]+)", re.I)
for f in sorted(os.listdir(d)):
    if not f.endswith(".bas"):
        continue
    text = open(os.path.join(d, f), encoding="utf-8").read()
    text = re.sub(r'"(?:[^"]|"")*"', "", text)
    text = re.sub(r"'.*", "", text)
    print(f, "->", sorted(set(pat.findall(text))))
