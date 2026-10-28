"""Попереднє наголошування тексту для озвучки (запускати, скажімо, раз на місяць або після великих правок).

Запуск у середовищі з ukrainian-word-stress (наприклад, D:\\zp-tts\\venv):
    D:\\zp-tts\\venv\\Scripts\\python.exe -X utf8 voice_mod\\make_stress.py --envfile D:\\zp-tts\\env.ps1

Читає text/translation.json, наголошує кожен рядок тексту подій, розмов, повідомлень і довідника
та пише в voice_mod/stress/:
  stressed_uk.tsv   — відбиток рядка, ключ, текст з наголосами («+» перед голосною). Генерується.
  stress_review.txt — слова, де бібліотека не визначилася (кілька наголосів): кандидати на правку.
  stress_fixes.txt  — РУЧНІ правки наголосів (по слову на рядок, «працівн+ики»). Скрипт його не перезаписує;
                      правки діють і тут, і в грі одразу, без перегенерації.
Повторний запуск наголошує лише нові й змінені рядки.
"""
import argparse
import collections
import json
import os
import re
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "mod_files"))
import stress_util as su  # noqa: E402

TABLES = ("NPCEvent", "NPCDialog", "OtherKeyMessage", "MachineHandbookEntry", "ContemplationType", "MachineProject",
          "CityTask")
FIELDS = ("display_name", "description", "high_text", "strategy_tip", "alt_description", "alt_strategy_tip",
          "reward_header_text")

ap = argparse.ArgumentParser()
ap.add_argument("--envfile")
ap.add_argument("--out-dir", default=os.path.join(HERE, "stress"))
ap.add_argument("--limit", type=int, default=0, help="наголосити не більше N нових рядків (для перевірки)")
args = ap.parse_args()
sys.stdout.reconfigure(encoding="utf-8")

if args.envfile and os.path.exists(args.envfile):
    for line in open(args.envfile, encoding="utf-8-sig"):
        m = re.match(r'\s*\$env:(\w+)\s*=\s*"([^"]*)"\s*$', line)
        if m and m.group(1) != "PATH":
            os.environ[m.group(1)] = m.group(2)

os.makedirs(args.out_dir, exist_ok=True)
tsv_path = os.path.join(args.out_dir, "stressed_uk.tsv")
fixes_path = os.path.join(args.out_dir, "stress_fixes.txt")
review_path = os.path.join(args.out_dir, "stress_review.txt")

if not os.path.exists(fixes_path):
    with open(fixes_path, "w", encoding="utf-8") as f:
        f.write("# Ручні правки наголосів для озвучки. По слову на рядок, наголос — «+» перед голосною:\n"
                "#   працівн+ики\n#   зам+ок        (діє на всі форми «замок» у тексті; для омографів обирайте частішу)\n")

data = json.load(open(os.path.join(os.path.dirname(HERE), "text", "translation.json"), encoding="utf-8"))
CHOICE_PARTS = ("choice", "option", "dialog_choice", "available_outcome")
CHOICE_FIELDS = ("display_name", "description", "strategy_tip", "line_1", "line_2", "alt_description", "alt_strategy_tip")


def wanted(p):
    if not p or p[0] not in TABLES:
        return False
    if len(p) == 3:
        return p[2] in FIELDS
    return len(p) >= 5 and p[2] in CHOICE_PARTS and p[-1] in CHOICE_FIELDS  # підказки до варіантів


items = {}  # відбиток -> (ключ, рядок)
for key, row in data.items():
    p = key.split("|")
    if not wanted(p):
        continue
    uk = row.get("Українська") or ""
    for line in uk.replace("\r", "").split("\n"):
        if line.strip():
            items.setdefault(su.line_hash(line), (key, line))

old = {}
if os.path.exists(tsv_path):
    for line in open(tsv_path, encoding="utf-8"):
        parts = line.rstrip("\n").split("\t")
        if len(parts) == 3:
            old[parts[0]] = parts[2]

# Заміни для голосу (ЛКМ → лівою кнопкою миші…). Якщо файл правил змінився — перенаголошуємо рядки, яких він стосується.
rules_path = os.path.join(args.out_dir, "speech_replace.tsv")
rules = su.load_speech_rules(rules_path)
sig_path = os.path.join(args.out_dir, ".speech_replace.sig")
sig = su.rules_signature(rules_path)
old_sig = open(sig_path, encoding="utf-8").read().strip() if os.path.exists(sig_path) else ""
if sig != old_sig:
    touched = [h for h, (_, line) in items.items() if h in old and any(rx.search(line) for rx, _ in rules)]
    for h in touched:
        old.pop(h, None)
    print(f"правила заміни змінились: перенаголошую {len(touched)} рядків")

todo = [h for h in items if h not in old]
if args.limit:
    todo = todo[:args.limit]
print(f"рядків: {len(items)}, уже наголошено: {len(items) - len([h for h in items if h not in old])}, нових: {len(todo)}")

done = dict((h, old[h]) for h in items if h in old)


def save():
    with open(tsv_path, "w", encoding="utf-8", newline="\n") as f:
        for h, (key, _) in sorted(items.items(), key=lambda kv: kv[1][0]):
            if h in done:
                f.write(f"{h}\t{key}\t{done[h]}\n")


if todo:
    from ukrainian_word_stress import Stressifier
    stressify = Stressifier()
    t0 = time.time()
    for i, h in enumerate(todo, 1):
        key, line = items[h]
        done[h] = su.acute_to_plus(stressify(su.speech_prepare(line, rules)))
        if i % 200 == 0 or i == len(todo):
            el = time.time() - t0
            print(f"  {i}/{len(todo)}  {el:.0f} с, лишилось ~{el / i * (len(todo) - i) / 60:.1f} хв", flush=True)
            save()
save()
with open(sig_path, "w", encoding="utf-8") as f:
    f.write(sig)

# Кандидати на ручну правку: слова з кількома наголосами (бібліотека не визначилась).
fixes = su.load_fixes(fixes_path)
amb = collections.Counter()
example = {}
for h, text in done.items():
    for w in su.WORD.findall(su.plus_to_acute(text)):
        if su.stress_count(w) >= 2 and su.strip_stress(w).lower() not in fixes:
            base = su.strip_stress(w).lower()
            amb[base] += 1
            example.setdefault(base, su.acute_to_plus(w))
with open(review_path, "w", encoding="utf-8") as f:
    f.write("# Слова з кількома наголосами — бібліотека не визначилась. Правильний варіант\n"
            "# допишіть у stress_fixes.txt (наприклад «р+оків»). Кількість — скільки разів трапляється.\n")
    for base, n in amb.most_common():
        f.write(f"{n}\t{example[base]}\n")
print(f"готово: {tsv_path}\nкандидатів на правку: {len(amb)} -> {review_path}")
