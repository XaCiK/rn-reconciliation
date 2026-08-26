# -*- coding: utf-8 -*-
import io
import os

file_path = r"C:\Users\o_antonov\NEW\RN\modReconcile.bas"
log_path = r"C:\Users\o_antonov\NEW\RN\verify_cp1251.log"

with io.open(file_path, "r", encoding="utf-8") as f:
    content = f.read()
with io.open(file_path, "wb") as f:
    f.write(content.encode("cp1251"))

with io.open(file_path, "rb") as f:
    raw = f.read()
read_back = raw.decode("cp1251")

tests = [
    u"\u0421\u043b\u043e\u0432\u0430\u0440\u044c",
    u"\u043d\u0435\u0444\u0442\u0435\u043f\u0440\u043e\u0434\u0443\u043a\u0442",
    u"\u0417\u0430\u0447\u0438\u0441\u043b\u0435\u043d\u043e",
    u"\u0413\u0421\u041c_\u0420\u041d\u041a\u0430\u0440\u0442",
    u"\u041a\u0430\u0440\u0442\u0430_\u0444\u0438\u043a\u0441\u0430\u0446\u0438\u044f",
    u"\u041d\u0435 \u0437\u0430\u0447\u0438\u0441\u043b\u0435\u043d\u043e",
    u"RN_Slovar",
    u"RN_Otchet_Bank",
]

lines = []
lines.append(u"File size: {0} bytes".format(os.path.getsize(file_path)))
bom = raw[:3]
lines.append(u"First 3 bytes: {0}".format(u",".join(unicode(b) for b in bom)))
all_ok = True
if bom == b"\xef\xbb\xbf":
    lines.append(u"WARNING: UTF-8 BOM detected")
    all_ok = False
else:
    lines.append(u"No UTF-8 BOM (good for CP1251)")

for t in tests:
    if t in read_back:
        lines.append(u"OK: " + t)
    else:
        lines.append(u"FAIL: " + t)
        all_ok = False

sample = u""
for ln in read_back.splitlines():
    if u"\u0421\u043b\u043e\u0432\u0430\u0440\u044c" in ln:
        sample = ln.strip()
        break
lines.append(u"Sample line: " + sample)
lines.append(u"Total lines: {0}".format(len(read_back.splitlines())))
lines.append(u"RESULT: CP1251 VALID" if all_ok else u"RESULT: CP1251 INVALID")

out = u"\n".join(lines)
with io.open(log_path, "w", encoding="utf-8") as f:
    f.write(out)
print(out.encode("utf-8"))
