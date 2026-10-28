"""UkrVoice: живий синтез голосом StyleTTS2 · RosKud (лише для локального тестування).

Мод гри запускає цей скрипт і пише в stdin рядки:
    SAY\t<стиль narr|inner|talk>\t<ключ>\t<текст UTF-8 у base64>
    STOP
Скрипт синтезує текст речення за реченням і грає через winsound; наступне речення
синтезується, поки звучить попереднє. Новий SAY або STOP обриває поточне.

Запуск вручну для перевірки:
    python tts_helper.py --model <тека StyleTTS2-RosKud> --envfile D:\\zp-tts\\env.ps1 --say "Привіт."
"""
import argparse
import base64
import os
import queue
import re
import sys
import threading
import time
import traceback
import unicodedata

HERE = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(HERE, "tts_helper_log.txt")
SR = 24000


def log(msg):
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(time.strftime("%H:%M:%S") + "\t" + msg + "\n")
    except Exception:
        pass


def load_envfile(path):
    """Підхоплює змінні з рядків виду $env:NAME = "value" (env.ps1 середовища zp-tts)."""
    if not path or not os.path.exists(path):
        return
    for line in open(path, encoding="utf-8-sig"):
        m = re.match(r'\s*\$env:(\w+)\s*=\s*"([^"]*)"\s*$', line)
        if m and m.group(1) != "PATH":
            os.environ[m.group(1)] = m.group(2)


ap = argparse.ArgumentParser()
ap.add_argument("--model", required=True)
ap.add_argument("--envfile")
ap.add_argument("--speed", type=float, default=1.0)
ap.add_argument("--lufs", type=float, default=-21.0, help="цільова гучність (ZERO PARADES: −24.6)")
ap.add_argument("--raw", action="store_true", help="без постобробки (для порівняння)")
ap.add_argument("--say", help="озвучити одну фразу й вийти (перевірка)")
args = ap.parse_args()
load_envfile(args.envfile)

import numpy as np  # noqa: E402
import torch  # noqa: E402
import soundfile as sf  # noqa: E402
import winsound  # noqa: E402
import librosa  # noqa: E402
import pyloudnorm as pyln  # noqa: E402
from ipa_uk import ipa  # noqa: E402
from ukrainian_word_stress import Stressifier  # noqa: E402
from styletts2_inference.models import StyleTTS2  # noqa: E402

sys.path.insert(0, HERE)
import stress_util as su  # noqa: E402

ACUTE = su.ACUTE
dev = "cuda" if torch.cuda.is_available() else "cpu"
t0 = time.time()
model = StyleTTS2(config_path=os.path.join(args.model, "config.yml"),
                  weights_path=os.path.join(args.model, "pytorch_model.bin"), device=dev)
styles = {v: torch.load(os.path.join(args.model, "voices", f"roskud_{v}.pt"), map_location=dev)
          for v in ("narr", "inner", "talk")}
stressify = Stressifier()
log(f"модель завантажено на {dev} за {time.time() - t0:.1f} с")

# Попередньо наголошений текст (make_stress.py) і ручні правки наголосів.
STRESSED = {}
_tsv = os.path.join(HERE, "stressed_uk.tsv")
if os.path.exists(_tsv):
    for _l in open(_tsv, encoding="utf-8"):
        _p = _l.rstrip("\n").split("\t")
        if len(_p) == 3:
            STRESSED[_p[0]] = _p[2]
FIXES_PATH = os.path.join(HERE, "stress_fixes.txt")
fixes_state = {"mtime": None, "fixes": {}}
log(f"наголошених рядків: {len(STRESSED)}")


def current_fixes():
    """Перечитує stress_fixes.txt, якщо його змінили, — правки діють без перезапуску гри."""
    try:
        m = os.path.getmtime(FIXES_PATH)
    except OSError:
        return {}
    if m != fixes_state["mtime"]:
        fixes_state["mtime"] = m
        fixes_state["fixes"] = su.load_fixes(FIXES_PATH)
        log(f"ручних правок наголосів: {len(fixes_state['fixes'])}")
    return fixes_state["fixes"]


RULES_PATH = os.path.join(HERE, "speech_replace.tsv")
rules_state = {"mtime": None, "rules": []}


def current_rules():
    try:
        m = os.path.getmtime(RULES_PATH)
    except OSError:
        return []
    if m != rules_state["mtime"]:
        rules_state["mtime"] = m
        rules_state["rules"] = su.load_speech_rules(RULES_PATH)
        log(f"замін для голосу: {len(rules_state['rules'])}")
    return rules_state["rules"]


stats = {"ready": 0, "live": 0}


def stressed_line(line):
    pre = STRESSED.get(su.line_hash(line))
    if pre is not None:
        stats["ready"] += 1
        s = su.plus_to_acute(pre)
    else:
        stats["live"] += 1
        s = stressify(su.speech_prepare(line, current_rules()))
    return su.apply_fixes(s, current_fixes())


def split_parts(line):
    parts, cur = [], ""
    for i, ch in enumerate(line):
        cur += ch
        if ch in ".?!:…" and i < len(line) - 1 and line[i + 1] == " " and len(cur) > 20:
            parts.append(cur)
            cur = ""
    if cur.strip():
        parts.append(cur)
    return parts


def text_parts(text):
    """Текст гри -> речення з уже розставленими наголосами."""
    for line in text.replace("\r", "").split("\n"):
        if line.strip():
            for part in split_parts(stressed_line(line)):
                yield part


def synth(part, voice):
    """Синтез речення, в якому вже стоять наголоси (U+0301)."""
    t = unicodedata.normalize("NFC", part.strip().replace('"', "").replace("«", "").replace("»", ""))
    t = re.sub(r"[‐‑‒–—―−]", "-", t)
    t = re.sub(r"\s+", " ", t).strip()
    if not t:
        return None
    if t[-1] not in ".?!:-…":
        t += "."
    t = re.sub(r" - ", ": ", t)
    ps = ipa(t)
    if not ps:
        return None
    with torch.no_grad():
        wav = model(model.tokenizer.encode(ps), speed=args.speed, s_prev=styles.get(voice, styles["narr"]))
    return wav.cpu().numpy()


OUT_SR = 48000
GAP = 0.12  # пауза між реченнями, с
meter = pyln.Meter(SR)


def post(y):
    """Як post.py у ZERO PARADES: обрізати тишу, вирівняти гучність, м'який лімітер, плавні краї;
    плюс якісний ресемплінг у 48 кГц, щоб Windows не перетворював частоту сам."""
    if args.raw:
        return y.astype(np.float32), SR
    y = y.astype(np.float32)
    _, (s, e) = librosa.effects.trim(y, top_db=40, frame_length=1024, hop_length=256)
    pad = int(0.05 * SR)
    y = y[max(0, s - pad):min(len(y), e + pad)]
    if len(y) >= int(0.4 * SR):
        loud = meter.integrated_loudness(y)
        if np.isfinite(loud):
            y = y * (10 ** ((args.lufs - loud) / 20))
    t, c = 0.5, 0.708  # піки вище −6 dBFS плавно до стелі −3 dBFS
    m = np.abs(y) > t
    if m.any():
        y[m] = np.sign(y[m]) * (t + (c - t) * np.tanh((np.abs(y[m]) - t) / (c - t)))
    f = int(0.01 * SR)
    if len(y) > 2 * f:
        y[:f] *= np.linspace(0, 1, f)
        y[-f:] *= np.linspace(1, 0, f)
    return librosa.resample(y, orig_sr=SR, target_sr=OUT_SR, res_type="soxr_hq").astype(np.float32), OUT_SR


requests = queue.Queue()
current = {"stop": threading.Event()}
# що зараз звучить: main — текст вікна (подія, розмова…), hover/hover_auto — підказка
playing = {"kind": None, "active": False, "until": 0.0}


def main_playing():
    return playing["kind"] == "main" and (playing["active"] or time.monotonic() < playing["until"])
tmp_files = [os.path.join(HERE, "_tts_a.wav"), os.path.join(HERE, "_tts_b.wav")]


def stop_playback():
    current["stop"].set()
    winsound.PlaySound(None, 0)


def worker():
    try:
        _worker()
    except Exception:
        log("потік синтезу впав\n" + traceback.format_exc())


def _worker():
    flip = 0
    while True:
        voice, key, text, stop, kind = requests.get()
        if stop.is_set():
            log(f"{key}: скасовано до початку")
            continue
        t_req = time.time()
        play_until = 0.0
        first = True
        playing.update(kind=kind, active=True, until=0.0)
        try:
            for part in text_parts(text):
                wav = synth(part, voice)
                if stop.is_set():
                    break
                if wav is None:
                    log(f"{key}: порожній синтез для «{part[:60]}»")
                    continue
                wav, sr = post(wav)
                path = tmp_files[flip]
                flip ^= 1
                sf.write(path, wav, sr, subtype="PCM_16")
                wait = play_until - time.monotonic()
                if wait > 0 and stop.wait(wait):
                    break
                winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC | winsound.SND_NODEFAULT)
                play_until = time.monotonic() + len(wav) / sr + GAP
                playing["until"] = play_until
                if first:
                    log(f"{key}: перший звук через {time.time() - t_req:.2f} с "
                        f"(наголоси: готових {stats['ready']}, на льоту {stats['live']})")
                    first = False
        except Exception:
            log("помилка синтезу " + key + "\n" + traceback.format_exc())
        if stop is current["stop"]:
            playing["active"] = False


def say(voice, key, text, kind="main"):
    stop_playback()
    current["stop"] = threading.Event()
    playing.update(kind=kind, active=True, until=0.0)
    requests.put((voice, key, text, current["stop"], kind))


threading.Thread(target=worker, daemon=True).start()
post(synth(next(text_parts("Готово, можна грати.")), "narr"))  # прогрів CUDA й обробки звуку, щоб перша репліка не чекала
log("готовий до роботи")

if args.say:
    say("narr", "test", args.say)
    time.sleep(1.0)
    while not requests.empty():
        time.sleep(0.2)
    time.sleep(8)
    sys.exit(0)

for raw in sys.stdin:
    line = raw.rstrip("\r\n")
    if not line:
        continue
    if line == "STOP":
        stop_playback()
        playing["active"] = False
        continue
    if line == "STOP_HOVER":  #підказка зникла: зупиняємо лише підказку, основний текст хай звучить
        if playing["kind"] in ("hover", "hover_auto"):
            stop_playback()
            playing["active"] = False
        continue
    p = line.split("\t")
    if p[0] == "SAY" and len(p) >= 4:
        try:
            text = base64.b64decode(p[3]).decode("utf-8")
        except Exception:
            log("не вдалося розкодувати: " + line[:80])
            continue
        kind = p[4] if len(p) >= 5 else "main"
        if kind == "hover_auto" and main_playing():
            log(f"SAY {p[2]}: пропущено — звучить основний текст (Ctrl прочитає підказку)")
            continue
        log(f"SAY {kind} {p[2]} ({len(text)} симв.)")
        say(p[1], p[2], text, kind)
log("stdin закрито, вихід")
