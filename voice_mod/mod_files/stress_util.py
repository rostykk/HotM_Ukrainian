"""Спільне для наголосів: формат «+ перед голосною» ↔ знак наголосу, ручні правки, відбиток рядка.

У файлах для людей наголос пишемо плюсом перед голосною: «зам+ок». Модель і ipa_uk чекають
комбінований гострий наголос (U+0301) після голосної: «замо́к».
"""
import hashlib
import re
import unicodedata

ACUTE = "́"
VOWELS = "аеєиіїоуюяАЕЄИІЇОУЮЯ"
LETTERS = "А-Яа-яЄєІіЇїҐґ"
WORD = re.compile("[" + LETTERS + "'’ʼ" + ACUTE + "+]+")


def line_hash(line):
    """Відбиток рядка тексту гри; однаковий у скрипті наголошування і в помічнику."""
    return hashlib.md5(unicodedata.normalize("NFC", line.replace("\r", "").strip()).encode("utf-8")).hexdigest()


def clean(text):
    """Прибирає те, що не читаємо вголос: теги, підстановки."""
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\{[^}]*\}", " ", text)
    return re.sub(r"[ \t]+", " ", text).strip()


def load_speech_rules(path):
    """speech_replace.tsv: «регулярний вираз<TAB>заміна» — лише для голосу (ЛКМ → лівою кнопкою миші тощо)."""
    rules = []
    try:
        lines = open(path, encoding="utf-8-sig").read().splitlines()
    except OSError:
        return rules
    for line in lines:
        if not line.strip() or line.lstrip().startswith("#") or "\t" not in line:
            continue
        pat, repl = line.split("\t", 1)
        try:
            rules.append((re.compile(pat), repl))
        except re.error:
            pass
    return rules


def speech_prepare(line, rules):
    """Рядок гри -> текст для синтезу (до наголошування): заміни для голосу, без тегів і підстановок."""
    for rx, repl in rules:
        line = rx.sub(repl, line)
    return clean(line)


def rules_signature(path):
    try:
        return hashlib.md5(open(path, "rb").read()).hexdigest()
    except OSError:
        return ""


def acute_to_plus(s):
    return re.sub("([" + VOWELS + "])" + ACUTE, r"+\1", s)


def plus_to_acute(s):
    return re.sub(r"\+([" + VOWELS + "])", "\\1" + ACUTE, s)


def strip_stress(word):
    return word.replace(ACUTE, "").replace("+", "")


def load_fixes(path):
    """stress_fixes.txt: по слову на рядок, «працівн+ики» або «працівники<TAB>працівн+ики». # — коментар."""
    fixes = {}
    try:
        lines = open(path, encoding="utf-8-sig").read().splitlines()
    except OSError:
        return fixes
    for line in lines:
        line = line.split("#", 1)[0].strip()
        if not line:
            continue
        parts = line.split()
        form = parts[-1]
        base = strip_stress(parts[0]).lower()
        if "+" in form:
            fixes[base] = form.lower()
    return fixes


def apply_fixes(stressed, fixes):
    """Замінює наголос у словах зі списку ручних правок (текст зі знаком U+0301)."""
    if not fixes:
        return stressed

    def repl(m):
        tok = m.group(0)
        base = strip_stress(tok)
        form = fixes.get(base.lower())
        if not form:
            return tok
        out = plus_to_acute(form)
        if base[:1].isupper():
            out = out[:1].upper() + out[1:]
        return out

    return WORD.sub(repl, stressed)


def stress_count(word):
    return word.count(ACUTE)
