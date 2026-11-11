import os

fp = r"D:\GG\A\nx_tieuhoc\build.ps1"
with open(fp, "r", encoding="utf-8") as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if "Copy-Item -Path" in l:
        lines[i] = "Copy-Item -Path $dataDir -Destination $dstData -Force -Recurse\n"
        print(f"Replaced line {i+1}")

with open(fp, "w", encoding="utf-8") as f:
    f.writelines(lines)

print("Updated build.ps1")
