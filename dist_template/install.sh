#!/bin/sh
# Українізатор Heart of the Machine — встановлення для Linux / Steam Deck.
# Використання: sh install.sh [шлях до теки HeartOfTheMachine]
HERE="$(cd "$(dirname "$0")" && pwd)"
PY="$(command -v python3 || command -v python)"
if [ -z "$PY" ]; then
  echo "Потрібен python3. Або встановіть вручну — див. ЧИТАЙМЕНЕ.txt"
  exit 1
fi
exec "$PY" - "$HERE" "${HOTM_UK_MODE:-install}" "$1" <<'PYEOF'
import glob, os, re, shutil, sys
here, mode, game = sys.argv[1], sys.argv[2], (sys.argv[3] if len(sys.argv) > 3 else "")
LT = os.path.join("GameData", "Configuration", "LocalizationType", "LocalizationTypes.xml")
ok = lambda p: bool(p) and os.path.isfile(os.path.join(p, LT))

def find():
    home = os.path.expanduser("~")
    roots = [os.path.join(home, ".local/share/Steam"), os.path.join(home, ".steam/steam"),
             os.path.join(home, ".var/app/com.valvesoftware.Steam/.local/share/Steam"),
             os.path.join(home, "Library/Application Support/Steam")]
    libs = list(roots)
    for r in roots:
        vdf = os.path.join(r, "steamapps", "libraryfolders.vdf")
        if os.path.isfile(vdf):
            libs += re.findall(r'"path"\s+"([^"]+)"', open(vdf, encoding="utf-8", errors="replace").read())
    libs += glob.glob("/run/media/*") + glob.glob("/run/media/*/*") + glob.glob("/media/*/*")
    cands = [here, os.path.dirname(here)] + [os.path.join(l, "steamapps", "common", "HeartOfTheMachine") for l in libs]
    return next((c for c in cands if ok(c)), None)

if not ok(game):
    game = find()
if not game:
    try:
        with open("/dev/tty") as tty:
            print("Вставте шлях до теки HeartOfTheMachine: ", end="", flush=True)
            game = tty.readline().strip().strip('"')
    except OSError:
        game = ""
if not ok(game):
    sys.exit(f"У теці «{game}» не знайдено гру.")
print("Гра:", game)
lt = os.path.join(game, LT)
text = open(lt, "rb").read().decode("utf-8-sig")
nl = "\r\n" if "\r\n" in text else "\n"
rx = re.compile(r'(\r?\n)*[ \t]*<localization_type\s+id="Ukrainian".*?/>', re.S)
xlsx = os.path.join(game, "GameData", "LocalizationImports", "Ukrainian.xlsx")
if mode == "install":
    shutil.copyfile(os.path.join(here, "Ukrainian.xlsx"), xlsx)
    print("Скопійовано Ukrainian.xlsx")
    block = open(os.path.join(here, "localization_type_uk.xml"), encoding="utf-8-sig").read().strip("\r\n")
    block = block.replace("\r\n", "\n").replace("\n", nl)
    m = rx.search(text)
    if m:
        text = text[:m.start()] + nl + nl + block + text[m.end():]
    else:
        shutil.copyfile(lt, lt + ".bak_before_ukrainian")
        i = text.rfind("</root>")
        text = text[:i].rstrip() + nl + nl + block + nl + nl + text[i:]
    print("Українську мову додано в LocalizationTypes.xml")
else:
    if os.path.exists(xlsx):
        os.remove(xlsx)
        print("Видалено Ukrainian.xlsx")
    text = rx.sub("", text, count=1)
    print("Українську мову прибрано з LocalizationTypes.xml")
with open(lt, "w", encoding="utf-8", newline="") as f:
    f.write(text)
print("Готово!" + (" Виберіть «Українська» у списку мов унизу головного меню." if mode == "install" else ""))
PYEOF
