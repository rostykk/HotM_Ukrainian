"""Пише en_lookup.tsv для прототипу: ключ локалізації -> англійський оригінал (переноси як \\n).

python -I -X utf8 voice_mod/make_en_lookup.py <куди_писати.tsv>
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import hotm_uk as h  # noqa: E402

TABLES = ("NPCEvent", "NPCDialog", "OtherKeyMessage", "MachineHandbookEntry", "ContemplationType")

out = sys.argv[1]
rows, _, _ = h.read_work(h.WORK_FILE)
n = 0
with open(out, "w", encoding="utf-8", newline="\n") as f:
    for r in rows:
        k = r[h.COL_KEY]
        if k.split("|")[0] not in TABLES:
            continue
        en = (r.get(h.COL_EN) or "").replace("\r", "").replace("\t", " ").replace("\n", "\\n")
        if en.strip():
            f.write(k + "\t" + en + "\n")
            n += 1
print("en_lookup:", n, "рядків ->", out)
