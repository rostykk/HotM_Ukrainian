# HotM_Ukrainian_work — для Claude

Фанатська українська локалізація Heart of the Machine разом із DLC Kinship: підтримуємо переклад, коли гра
змінюється. Спілкування з користувачем, код, повідомлення скрипта й коміти — українською. Повний процес
описано в `README.md`, стиль перекладу — у `STYLE_uk.md`. Перед роботою з текстом прочитай обидва.

## Файли

- `hotm_uk.py` — увесь конвеєр (`-h` показує команди).
- `uk_translation.xlsx` — **головний** робочий файл зі статусами рядків: `new`, `changed`, `auto`, `mt`
  (машинний переклад, ще не вичитаний), `ok`. На 13.10.2026: 67 рядків `ok`, решта `mt`.
- `glossary.xlsx` — глосарій.
- `text/*.json` — текстова копія обох xlsx для git diff. Генерує скрипт, вручну не правити.
- `snapshots/en_<версія>_b<збірка Steam>.json` — англійський текст кожної збірки. Реліз 1.0 зроблено
  з `en_2_001_Update89.json`.
- `batches/`, `backup/`, `changes/`, `dist/` — генеровані, в git не йдуть.
- `.claude/skills/ai-translate` — скіл перекладу рядків `new`/`changed` через ШІ-батчі.
- Гра лежить у `D:\SteamLibrary\steamapps\common\HeartOfTheMachine`. На машині автора проєкт — це тека
  `UkrainianLocalization` всередині гри. У копії друзів поруч лежить `..\HotM_Ukrainian_project` —
  оригінал автора, **не змінювати**.

## Запуск

```
$env:PYTHONIOENCODING="utf-8"
.venv\Scripts\python.exe hotm_uk.py --game "D:\SteamLibrary\steamapps\common\HeartOfTheMachine" <команда>
```

- Залежності встановлено лише в `.venv` (див. `requirements.txt`). Глобальний `python` скрипта не запустить.
  Якщо `.venv` немає, створи його: `python -m venv .venv`, потім `.venv\Scripts\python.exe -m pip install -r requirements.txt`.
- `sync --dry-run` нічого не записує. Справжній `sync` переписує `uk_translation.xlsx`, а `build` пише в теку
  гри.
- Після правок xlsx в обхід скрипта запусти `dump` або `build`, щоб оновилась `text/`.

## Якість мови

- `russisms_uk.txt` — русизми й кальки: без «?» — помилка (`check` блокує), з «?» — підозра (примітка;
  `build` пише `build_suspects.txt`). Побачив новий русизм — додай туди рядок.
- `export --style` — батчі для літературного редактора (природність, кальки, витієватість); відповідь
  містить лише змінені id (`"partial": true`). Правила — розділ «Природність» у `STYLE_uk.md`.
- Автор дуже чутливий до русизмів і «витієватого» тексту: перевіряй фрази вголос, пиши просто.

## Після оновлення гри

`sync` → коміт `sync: <версія> b<збірка>` → переклад рядків `new`/`changed` (скіл `ai-translate`) →
`build` → коміт. Перед `dist --release X` — новий запис зверху `dist_template/changelog.txt` (він іде в
посібник Steam, ЧИТАЙМЕНЕ і `dist/steam_changenote_X.txt`).

## Факти, які вже перевірено (не досліджувати заново)

- Англійський оригінал — XML у `GameData\Configuration` і `Expansions\*`. Файл `English.xlsx` і тека
  `LocalizationExports` у грі застаріли.
- Ключі мають вигляд `Таблиця|id|поле` (для вкладених вузлів — `…|вузол|id|поле`); рядки з `0_Language` — просто `id`.
- `ALL-LOCA-LINKS.txt` (роздільник `中`): ключ ліворуч бере переклад ключа праворуч.
- NPCEvent з `tl="OC"` зберігаються під `ContemplationType|<id без Cont_>|…`. Якщо ціль редиректу —
  `NPCEvent|Cont_*`, `build` дописує переклад і під цим ключем.
- Arcen змінює тексти, не змінюючи id версії. Збірку визначає `buildid` у
  `steamapps\appmanifest_2001070.acf`.
- Steam перезаписує `LocalizationTypes.xml`, і `build` повертає туди блок мови з `localization_type_uk.xml`.

## Git

- Тег `v1.0` — незмінений стан проєкту автора. Xlsx комітиться разом з оновленою `text/`.
- `.gitattributes` вимикає перетворення кінців рядків: `.sh` має лишатися з LF, `.bat` — з CRLF.
- Офіційні арти Arcen у `dist_template/art_src/` треба прибрати, перш ніж репозиторій стане публічним.
