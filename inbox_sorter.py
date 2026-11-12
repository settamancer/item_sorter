#!/usr/bin/env python3
"""
Раскладывает новые файлы из папки «Входящие» по папкам номенклатуры фотобанка.

Источник правды - книга photobank_rules.xlsx (создаётся командой --init):
  Настройки         пути, режим (тест/боевой)
  Категории         номер -> папка категории
  Новые артикулы    артикул -> номер категории (для артикулов, которых ещё нет)
  Папки             структура папки номенклатуры: что создавать, где транслитерировать
                    имена и дописывать _<артикул>
  Правила           куда класть файл по ключевым словам / маске / расширению
  Серии             серия -> список артикулов
  Общие документы   сертификат / отказное письмо -> серия (копия в каждый артикул)

Для каждого файла во «Входящих»:
  1. Файл есть в «Общих документах» - копия в каждый артикул серии.
  2. Иначе артикул берётся из имени подпапки «Входящих» (Входящие\\05-0900\\...) или из
     имени файла (05-0900_T1.png). Существующая папка артикула (в любой категории)
     дополняется; если её нет, но артикул указан в «Новых артикулах», она создаётся.
  3. Внутри папки файл раскладывается по «Правилам»; имя транслитерируется и
     дополняется _<артикул> согласно листу «Папки».
  4. Если файл с таким именем уже есть и отличается - старый переносится в архив.
  5. Нераспознанные файлы попадают в отчёт (в боевом режиме - в «_Не разобрано»).

Режим «тест»: файлы копируются, «Входящие» не меняются.
Режим «боевой»: разложенные файлы удаляются из «Входящих».
Отчёт (журнал, не разобрано, полнота) пишется в Отчёт.xlsx, журнал копится в журнал.csv.
"""

import argparse
import csv
import datetime
import fnmatch
import hashlib
import os
import re
import shutil
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

import photobank_sync as ps

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    print("Нужна библиотека openpyxl. Установите: python -m pip install openpyxl")
    sys.exit(1)

BOOK_DEFAULT = "photobank_rules.xlsx"
ROOT_LABEL = "(корень)"
SERVICE = ps.SERVICE
OLD_PHOTOS_TARGET = f"{SERVICE}/Старые фото"
OLD_PHOTOS_DIR = "старые фото"
# Похоже на артикул (для подсказки в отчёте, когда артикул не найден).
ARTICLE_HINT_RE = re.compile(r"(?<![0-9A-Za-zА-Яа-яЁё-])\d{2}-[0-9A-Za-zА-Яа-яЁё]{3,}(?:-[0-9A-Za-zА-Яа-яЁё]+)*")

S_INFO, S_SETTINGS, S_CATS, S_NEW, S_FOLDERS, S_RULES, S_SERIES, S_SHARED = (
    "Инструкция", "Настройки", "Категории", "Новые артикулы", "Папки", "Правила", "Серии", "Общие документы")


def log(msg):
    print(time.strftime("[%H:%M:%S] ") + msg, flush=True)


norm = ps.norm


def yes(v):
    return norm(str(v or "")) in {"да", "yes", "1", "true", "истина", "+"}


def split_list(v):
    return [x.strip() for x in re.split(r"[;\n]", str(v or "")) if x.strip()]


# --------------------------------------------------------------------------- шаблон книги

INFO_TEXT = [
    "Книга правил для inbox_sorter.py - источник правды для раскладки новых файлов.",
    "",
    "Как добавить файлы:",
    "  1. Положите файлы во «Входящие» (путь - на листе «Настройки»).",
    "     Файлы одного артикула можно положить в подпапку с именем артикула: Входящие\\05-0900\\...",
    "     Иначе артикул берётся из имени файла (05-0900_T1.png, Паспорт 05-0900.pdf).",
    "  2. Для нового артикула (его папки ещё нет) добавьте строку на лист «Новые артикулы».",
    "  3. Общий сертификат / отказное письмо: строка на листе «Общие документы» + серия на листе «Серии».",
    "  4. Запустите run_inbox.bat. Результат - в Отчёт.xlsx.",
    "",
    "Режим «тест» только копирует файлы; «боевой» удаляет разложенные файлы из «Входящих».",
    "",
    "Правила (лист «Правила») проверяются по порядку, срабатывает первое подходящее.",
    "  Условие - через «;»: слово (ищется в имени без учёта регистра), маска с * и ? (паспорт*),",
    "  или регулярное выражение с префиксом re:. {артикул} подставляет артикул.",
    "  Расширения - через «;», пусто = любые. Папка - пусто или (корень) = корень папки номенклатуры.",
    "  Файлы, уже лежащие в стандартной подпапке (Входящие\\05-0900\\Сертификаты\\...), остаются в ней.",
    "  Любая папка «Старые фото» переносится целиком в «Служебные материалы/Старые фото».",
    "",
    "Префикс (лист «Папки») дописывается в начало имени каждого файла этой папки, если имя",
    "  ещё не начинается с него. Пусто = без префикса. {артикул} подставляет артикул: {артикул}_",
    "  В папках с транслитерацией префикс тоже транслитерируется.",
]

SETTINGS_ROWS = [
    ("Входящие", "Входящие", "Куда вручную кладут новые файлы. Относительный путь - от папки с этой книгой."),
    ("Назначение", "Фотобанк", "Корень фотобанка (<категория>\\<артикул>). Позже - путь на Synology: "
                               r"\\192.168.13.200\дм\ОПМ\БАЗА ПО ПРОДУКТАМ - ФОТОБАНК"),
    ("Режим", "тест", "тест - только копировать; боевой - удалять разложенные файлы из «Входящих», "
                      "нераспознанные переносить в «Не разобрано»"),
    ("Не разобрано", "_Не разобрано", "Подпапка «Входящих» для нераспознанных файлов (боевой режим)"),
    ("Архив", f"{SERVICE}/Архив", "Куда в папке номенклатуры уходит старая версия файла при замене"),
    ("Отчёт", "Отчёт.xlsx", "Отчёт последнего запуска (перезаписывается)"),
    ("Журнал", "журнал.csv", "Накопительный журнал всех действий"),
    ("Проверять полноту", "да", "Лист «Полнота» в отчёте: у каких артикулов нет паспорта, фото T1, сертификатов"),
]

FOLDER_ROWS = [  # папка, создавать всегда, транслитерация, дописывать артикул, префикс
    (ROOT_LABEL, "да", "нет", "да", ""),
    ("Описание", "да", "да", "да", ""),
    ("Сертификаты", "да", "да", "да", ""),
    (SERVICE, "да", "нет", "да", ""),
    (f"{SERVICE}/Исходники", "нет", "нет", "нет", ""),
    (f"{SERVICE}/Протокол испытаний", "да", "нет", "да", ""),
    (f"{SERVICE}/Старые фото", "нет", "нет", "нет", ""),
    (f"{SERVICE}/Упаковка", "да", "нет", "да", ""),
    (f"{SERVICE}/Архив", "нет", "нет", "нет", ""),
    ("Фото товара", "да", "да", "да", ""),
]

IMG = "png; jpg; jpeg; webp; gif; bmp; tif; tiff"
RAW = "cr2; cr3; nef; arw; orf; raf; rw2; dng; psd; psb; tif; tiff; ai; eps; indd"
RULE_ROWS = [  # порядок, условие, расширения, папка, комментарий
    (10, r"re:^{артикул}_(?:[a-zа-яё]+\d*|3[dд]\d*)\.[^.]+$", IMG, "Фото товара",
     "Фото товара: 05-0900_T1, _Y1, _3D (в т.ч. с русскими Т/У/Д)"),
    (20, r"re:data[\s_.-]*sheet; даташит", "", SERVICE, "Datasheet - в корень служебных материалов"),
    (30, "паспорт*", "pdf", ROOT_LABEL, "Паспорт PDF - в корень"),
    (40, "паспорт*", "", SERVICE, "Паспорт в других форматах (docx) - в служебные"),
    (50, "сертификат; декларац; отказ; отрицательн; выписк; реестр; свидетельств", "", "Сертификаты",
     "Сертификаты, декларации, отказные письма, выписки ТОРП"),
    (60, "протокол", "", f"{SERVICE}/Протокол испытаний", ""),
    (70, "описани", "", "Описание", ""),
    (80, "этикет; упаков; коробк", "", f"{SERVICE}/Упаковка", ""),
    (90, "сплеш; splash", "", SERVICE, ""),
    (100, "", RAW, f"{SERVICE}/Исходники", "Исходники: RAW, PSD, TIFF"),
]


def series_from_structure_csv(path):
    """Серии и общие документы из старого structure.csv (код модели FO-R, FO-RM ...)."""
    series, shared = {}, {}
    if not path.is_file():
        return series, shared
    rows = list(csv.reader(path.read_text(encoding="utf-8-sig").splitlines(), delimiter=";"))
    for r in rows[1:]:
        if not r or not r[0].strip() or not r[-1].strip():
            continue
        code = Path(r[-1].strip()).stem
        series.setdefault(code, []).append(r[0].strip())
        for f in r[4:14]:
            f = f.strip()
            if f and not norm(f).startswith("паспорт"):
                shared.setdefault(f, code)
    return series, shared


def style_sheet(ws, widths):
    head = PatternFill("solid", fgColor="DDE7F3")
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = head
        c.alignment = Alignment(vertical="top", wrap_text=True)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions


def create_book(path, structure_csv):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = S_INFO
    for line in INFO_TEXT:
        ws.append([line])
    ws.column_dimensions["A"].width = 120
    ws["A1"].font = Font(bold=True, size=12)

    def sheet(name, headers, rows, widths):
        s = wb.create_sheet(name)
        s.append(headers)
        for r in rows:
            s.append(list(r))
        style_sheet(s, widths)
        return s

    sheet(S_SETTINGS, ["Параметр", "Значение", "Пояснение"], SETTINGS_ROWS, [22, 45, 110])
    sheet(S_CATS, ["№", "Папка категории"],
          [(n, f"{n} {name}") for n, name in ps.CATEGORIES.items()], [6, 60])
    sheet(S_NEW, ["Артикул", "№ категории", "Комментарий"], [], [18, 14, 60])
    sheet(S_FOLDERS, ["Папка", "Создавать всегда", "Транслитерация", "Дописывать артикул", "Префикс"],
          FOLDER_ROWS, [42, 18, 16, 20, 20])
    sheet(S_RULES, ["Порядок", "Условие", "Расширения", "Папка", "Комментарий"], RULE_ROWS,
          [10, 60, 40, 38, 55])
    series, shared = series_from_structure_csv(structure_csv)
    sheet(S_SERIES, ["Серия", "Артикулы", "Комментарий"],
          [(k, "; ".join(v), "из structure.csv") for k, v in series.items()], [16, 120, 20])
    sheet(S_SHARED, ["Файл (имя или маска)", "Серия", "Папка", "Комментарий"],
          [(f, s, "Сертификаты", "из structure.csv") for f, s in shared.items()], [90, 14, 18, 20])
    wb.save(path)


# --------------------------------------------------------------------------- чтение книги

@dataclass
class Rule:
    order: float
    conds: list
    exts: set
    folder: str


@dataclass
class Config:
    base: Path
    inbox: Path
    dest: Path
    combat: bool
    unsorted: str
    archive: str
    report: Path
    journal: Path
    completeness: bool
    categories: dict = field(default_factory=dict)      # № -> имя папки
    new_articles: dict = field(default_factory=dict)    # norm(артикул) -> (артикул, №)
    folders: list = field(default_factory=list)         # (путь, создавать, транслит, артикул, префикс)
    rules: list = field(default_factory=list)
    series: dict = field(default_factory=dict)          # norm(серия) -> [артикулы/маски]
    shared: list = field(default_factory=list)          # (маска, серия, папка)


def rows_of(wb, name):
    if name not in wb.sheetnames:
        raise ValueError(f"В книге нет листа «{name}»")
    ws = wb[name]
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        r = ["" if v is None else str(v).strip() if not isinstance(v, (int, float)) else v for v in r]
        if any(str(v).strip() for v in r):
            out.append(r + [""] * 6)
    return out


def folder_key(p):
    p = str(p).strip().replace("\\", "/").strip("/")
    return "" if norm(p) in {"", norm(ROOT_LABEL)} else p


def load_config(book):
    wb = openpyxl.load_workbook(book, read_only=True, data_only=True)
    base = book.parent
    st = {norm(r[0]): str(r[1]).strip() for r in rows_of(wb, S_SETTINGS)}

    def path_of(key, default):
        v = st.get(norm(key)) or default
        p = Path(v)
        return p if p.is_absolute() or v.startswith("\\\\") else base / p

    cfg = Config(
        base=base,
        inbox=path_of("Входящие", "Входящие"),
        dest=path_of("Назначение", "Фотобанк"),
        combat=norm(st.get("режим", "тест")).startswith("боев"),
        unsorted=st.get("не разобрано") or "_Не разобрано",
        archive=folder_key(st.get("архив") or f"{SERVICE}/Архив"),
        report=path_of("Отчёт", "Отчёт.xlsx"),
        journal=path_of("Журнал", "журнал.csv"),
        completeness=yes(st.get("проверять полноту", "да")),
    )
    for r in rows_of(wb, S_CATS):
        try:
            cfg.categories[int(float(r[0]))] = str(r[1]).strip()
        except (TypeError, ValueError):
            raise ValueError(f"Лист «{S_CATS}»: неверный номер категории «{r[0]}»")
    for r in rows_of(wb, S_NEW):
        art = str(r[0]).strip()
        try:
            num = int(float(r[1]))
        except (TypeError, ValueError):
            raise ValueError(f"Лист «{S_NEW}»: у артикула {art} не указан номер категории")
        if num not in cfg.categories:
            raise ValueError(f"Лист «{S_NEW}»: категории {num} нет на листе «{S_CATS}»")
        cfg.new_articles[norm(art)] = (art, num)
    for r in rows_of(wb, S_FOLDERS):
        cfg.folders.append((folder_key(r[0]), yes(r[1]), yes(r[2]), yes(r[3]), str(r[4]).strip()))
    if not any(f[0] == "" for f in cfg.folders):
        cfg.folders.append(("", True, False, True, ""))
    for r in rows_of(wb, S_RULES):
        try:
            order = float(r[0])
        except (TypeError, ValueError):
            order = 1e9
        exts = {e.lower().lstrip(".") for e in split_list(r[2])}
        cfg.rules.append(Rule(order, split_list(r[1]), exts, folder_key(r[3])))
    cfg.rules.sort(key=lambda x: x.order)
    for r in rows_of(wb, S_SERIES):
        cfg.series[norm(r[0])] = split_list(r[1])
    for r in rows_of(wb, S_SHARED):
        cfg.shared.append((str(r[0]).strip(), str(r[1]).strip(), folder_key(r[2]) or "Сертификаты"))
    wb.close()
    return cfg


# --------------------------------------------------------------------------- классификация

def rule_matches(rule, name, article):
    n = norm(name)
    ext = Path(name).suffix.lower().lstrip(".")
    if rule.exts and ext not in rule.exts:
        return False
    if not rule.conds:
        return True
    a = norm(article)
    for c in rule.conds:
        if c.lower().startswith("re:"):
            pat = c[3:].strip().replace("{артикул}", re.escape(a))
            if re.search(pat, n, re.IGNORECASE):
                return True
            continue
        c = norm(c.replace("{артикул}", a))
        if "*" in c or "?" in c:
            if fnmatch.fnmatchcase(n, c):
                return True
        elif c in n:
            return True
    return False


def canon_folders(cfg):
    """norm-кортеж частей пути -> каноническое написание (для стандартных папок)."""
    return {tuple(norm(p) for p in f[0].split("/")): f[0] for f in cfg.folders if f[0]}


def classify(cfg, rel_parts, article):
    """Относительный путь в папке номенклатуры для файла rel_parts (путь внутри «Входящих»
    после папки артикула, если она была)."""
    name, folders = rel_parts[-1], rel_parts[:-1]
    nf = tuple(norm(f) for f in folders)

    if OLD_PHOTOS_DIR in nf:
        k = nf.index(OLD_PHOTOS_DIR)
        return Path(OLD_PHOTOS_TARGET, *folders[k + 1:], name)

    canon = canon_folders(cfg)
    for k in range(len(nf), 0, -1):
        if nf[:k] in canon:
            return Path(canon[nf[:k]], *folders[k:], name)

    for rule in cfg.rules:
        if rule_matches(rule, name, article):
            return Path(rule.folder, name) if rule.folder else Path(name)
    if not folders:
        return Path(name)
    return Path(SERVICE, *folders, name)


def folder_flags(cfg, parent):
    """Самая длинная запись листа «Папки», в которую входит parent."""
    p = norm(parent.as_posix()) if str(parent) not in ("", ".") else ""
    best = None
    for f in cfg.folders:
        key = norm(f[0])
        if key == "" or p == key or p.startswith(key + "/"):
            if best is None or len(key) > len(norm(best[0])):
                best = f
    return best


def final_name(cfg, rel, article):
    """Префикс, транслитерация и _<артикул> по флагам листа «Папки»."""
    entry, _, translit, suffix, prefix = folder_flags(cfg, rel.parent)
    entry_parts = len(Path(entry).parts) if entry else 0
    below = list(rel.parent.parts[entry_parts:]) if str(rel.parent) not in ("", ".") else []
    name = rel.name
    prefix = prefix.replace("{артикул}", article)
    if prefix and not norm(name).startswith(norm(prefix)):
        name = prefix + name
    has_article = ps.article_pattern(article).search(norm(name)) is not None
    art = article
    if translit:
        below = [ps.to_latin(p) for p in below]
        name = ps.to_latin(name)
        art = ps.to_latin(article)
    if suffix and not has_article:
        p = Path(name)
        stem, ext = (p.stem, p.suffix) if p.suffix else (name, "")
        name = f"{stem}_{art}{ext}"
    head = Path(entry) if entry else Path()
    return head.joinpath(*below, name)


# --------------------------------------------------------------------------- фотобанк

@dataclass
class Bank:
    cats: dict = field(default_factory=dict)    # № -> Path существующей папки категории
    items: dict = field(default_factory=dict)   # norm(артикул) -> (артикул, Path, №)
    dups: list = field(default_factory=list)


def scan_bank(cfg):
    bank = Bank()
    if not os.path.isdir(ps.long_path(cfg.dest)):
        return bank
    dirs, _ = ps.list_dir(cfg.dest)
    for d in dirs:
        m = ps.CATEGORY_RE.match(d)
        if not m or int(m.group(1)) not in cfg.categories:
            continue
        num = int(m.group(1))
        bank.cats.setdefault(num, cfg.dest / d)
        items, _ = ps.list_dir(cfg.dest / d)
        for it in items:
            key = norm(it)
            if key in bank.items:
                bank.dups.append((it, bank.items[key][1], cfg.dest / d / it))
                continue
            bank.items[key] = (it, cfg.dest / d / it, num)
    return bank


class Resolver:
    """Находит артикул в пути файла и папку номенклатуры для него."""

    def __init__(self, cfg, bank):
        self.cfg, self.bank = cfg, bank
        self.known = {}
        for key, (art, num) in cfg.new_articles.items():
            self.known[key] = art
        for key, (art, _, _) in bank.items.items():
            self.known[key] = art
        self.patterns = {k: ps.article_pattern(a) for k, a in self.known.items()}

    def refresh(self, key, art):
        self.known[key] = art
        self.patterns[key] = ps.article_pattern(art)

    def from_path(self, rel):
        """(артикул, путь внутри папки артикула, []) или (None, причина, [артикулы для отчёта])."""
        for i, part in enumerate(rel[:-1]):
            if norm(part) in self.known:
                return self.known[norm(part)], rel[i + 1:], []
        for part in rel[:-1]:
            if ARTICLE_HINT_RE.fullmatch(part.strip()):
                return (None, f"папка «{part}»: такого артикула нет в фотобанке и на листе «{S_NEW}»",
                        [part.strip()])
        name = norm(rel[-1])
        hits = [k for k, pat in self.patterns.items() if pat.search(name)]
        if not hits:
            hint = ARTICLE_HINT_RE.search(rel[-1])
            if hint:
                return None, (f"артикул {hint.group(0)} не найден в фотобанке; "
                              f"добавьте его на лист «{S_NEW}» с номером категории"), [hint.group(0)]
            return None, "в имени файла нет артикула (положите файл в подпапку с именем артикула)", []
        best = max(hits, key=len)
        if any(h not in best for h in hits):
            arts = [self.known[h] for h in hits]
            return None, "в имени несколько разных артикулов: " + ", ".join(arts), arts
        return self.known[best], rel, []

    def item_dir(self, article, dry, created):
        """Папка номенклатуры: существующая или новая (по листу «Новые артикулы»)."""
        key = norm(article)
        if key in self.bank.items:
            return self.bank.items[key][1], None
        if key not in self.cfg.new_articles:
            return None, f"артикул {article} не найден в фотобанке и на листе «{S_NEW}»"
        art, num = self.cfg.new_articles[key]
        cat = self.bank.cats.get(num) or self.cfg.dest / self.cfg.categories[num]
        item = cat / art
        if not dry:
            for f, create, *_ in self.cfg.folders:
                if create:
                    os.makedirs(ps.long_path(item / f if f else item), exist_ok=True)
        self.bank.cats.setdefault(num, cat)
        self.bank.items[key] = (art, item, num)
        created.append(art)
        log(f"  создана папка номенклатуры: {cat.name}/{art}")
        return item, None


# --------------------------------------------------------------------------- копирование

def file_hash(p):
    h = hashlib.sha256()
    with open(ps.long_path(p), "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def same_file(a, b):
    sa, sb = os.stat(ps.long_path(a)), os.stat(ps.long_path(b))
    return sa.st_size == sb.st_size and file_hash(a) == file_hash(b)


def unique_path(p):
    if not os.path.exists(ps.long_path(p)):
        return p
    i = 2
    while True:
        c = p.with_name(f"{p.stem} ({i}){p.suffix}")
        if not os.path.exists(ps.long_path(c)):
            return c
        i += 1


class Run:
    def __init__(self, cfg, dry):
        self.cfg, self.dry = cfg, dry
        self.journal, self.unsorted, self.created = [], [], []
        self.unsorted_files = []   # файлы (в self.unsorted - строка на каждый артикул)
        self.written = {}   # norm(dst) -> исходный файл (защита от двух файлов в одно место)
        self.stamp = datetime.datetime.now()

    def record(self, src_rel, article, dst, action, note=""):
        self.journal.append({
            "Время": self.stamp.strftime("%Y-%m-%d %H:%M:%S"),
            "Файл во «Входящих»": src_rel,
            "Артикул": article or "",
            "Куда": str(dst.relative_to(self.cfg.dest)) if dst else "",
            "Действие": action,
            "Примечание": note,
        })

    def place(self, src, src_rel, article, item, rel_target):
        dst = item / rel_target
        key = norm(str(dst))
        if key in self.written and self.written[key] != src_rel:
            dst = unique_path(dst) if not self.dry else dst.with_name(f"{dst.stem} (2){dst.suffix}")
        self.written[norm(str(dst))] = src_rel
        exists = os.path.exists(ps.long_path(dst))
        if exists and same_file(src, dst):
            self.record(src_rel, article, dst, "уже есть")
            return True
        action, note = "добавлен", ""
        if not self.dry:
            os.makedirs(ps.long_path(dst.parent), exist_ok=True)
        if exists:
            stamp = self.stamp.strftime("%Y-%m-%d_%H%M%S")
            arch = item / self.cfg.archive / f"{dst.stem}_{stamp}{dst.suffix}"
            if not self.dry:
                arch = unique_path(arch)
                os.makedirs(ps.long_path(arch.parent), exist_ok=True)
                shutil.move(ps.long_path(dst), ps.long_path(arch))
            action, note = "заменён", f"старая версия: {arch.relative_to(item).as_posix()}"
        if not self.dry:
            shutil.copy2(ps.long_path(src), ps.long_path(dst))
        self.record(src_rel, article, dst, action, note)
        log(f"  {action}: {dst.relative_to(self.cfg.dest).as_posix()}  (из {src_rel})")
        return True

    def not_sorted(self, src_rel, reason, articles=()):
        """Строка «Не разобрано» на каждый не найденный артикул (или одна, если артикула нет)."""
        if isinstance(articles, str):
            articles = [articles]
        self.unsorted_files.append(src_rel)
        for art in articles or [""]:
            self.unsorted.append({"Файл во «Входящих»": src_rel, "Артикул": art, "Причина": reason})
            self.record(src_rel, art, None, "не разобран", reason)
        log(f"  НЕ РАЗОБРАН: {src_rel} - {reason}")


def match_shared(cfg, name):
    n = norm(name)
    for pat, series, folder in cfg.shared:
        p = norm(pat)
        if ("*" in p or "?" in p) and fnmatch.fnmatchcase(n, p) or n == p:
            return series, folder
    return None


def expand_series(cfg, resolver, series):
    items = cfg.series.get(norm(series))
    if items is None:
        return None
    out = []
    for a in items:
        if "*" in a or "?" in a:
            out += sorted(resolver.known[k] for k in resolver.known if fnmatch.fnmatchcase(k, norm(a)))
        else:
            out.append(a)
    return out


def process_file(run, resolver, src, rel):
    cfg = run.cfg
    src_rel = "/".join(rel)

    shared = match_shared(cfg, rel[-1])
    if shared:
        series, folder = shared
        articles = expand_series(cfg, resolver, series)
        if articles is None:
            run.not_sorted(src_rel, f"серии «{series}» нет на листе «{S_SERIES}»")
            return False
        missing = []
        for art in articles:
            item, err = resolver.item_dir(art, run.dry, run.created)
            if err:
                missing.append(art)
                continue
            run.place(src, src_rel, art, item, final_name(cfg, Path(folder, rel[-1]), art))
        if missing:
            run.not_sorted(src_rel, f"общий документ серии {series}: нет папки артикула; добавьте "
                                    f"его на лист «{S_NEW}» или уберите из серии "
                                    f"(не хватает папок: {len(missing)} из {len(articles)})", missing)
            return False
        return True

    article, inner, hints = resolver.from_path(rel)
    if article is None:
        run.not_sorted(src_rel, inner, hints)
        return False
    item, err = resolver.item_dir(article, run.dry, run.created)
    if err:
        run.not_sorted(src_rel, err, article)
        return False
    run.place(src, src_rel, article, item, final_name(cfg, classify(cfg, inner, article), article))
    return True


def inbox_files(cfg):
    unsorted_key = norm(cfg.unsorted)
    for rel in ps.walk_files(cfg.inbox):
        if rel and norm(rel[0]) == unsorted_key:
            continue
        yield rel


def prune_empty_dirs(root, keep):
    for dirpath, dirnames, filenames in os.walk(ps.long_path(root), topdown=False):
        p = Path(dirpath)
        if p == Path(ps.long_path(root)) or norm(p.name) == norm(keep):
            continue
        try:
            os.rmdir(dirpath)  # удаляется только пустая папка
        except OSError:
            pass


# --------------------------------------------------------------------------- отчёт

def completeness(cfg, bank):
    rows = []
    for key, (art, item, num) in sorted(bank.items.items(), key=lambda x: (x[1][2], x[0])):
        try:
            _, root_files = ps.list_dir(item)
        except OSError:
            continue

        def files_in(sub):
            try:
                return ps.list_dir(item / sub)[1]
            except OSError:
                return []

        photos, certs, desc = files_in("Фото товара"), files_in("Сертификаты"), files_in("Описание")
        passport = any(norm(f).startswith("паспорт") and f.lower().endswith(".pdf") for f in root_files)
        t1 = any(norm(Path(f).stem).endswith("_t1") for f in photos)
        lack = [n for n, ok in (("паспорт", passport), ("фото T1", t1), ("сертификаты", certs)) if not ok]
        rows.append({
            "Категория": bank.cats.get(num, Path(str(num))).name, "Артикул": art,
            "Паспорт PDF": "да" if passport else "нет", "Фото T1": "да" if t1 else "нет",
            "Фото товара, шт": len(photos), "Сертификаты, шт": len(certs), "Описание, шт": len(desc),
            "Не хватает": ", ".join(lack),
        })
    return rows


def write_report(cfg, run, bank, dry):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Сводка"
    counts = {}
    for j in run.journal:
        counts[j["Действие"]] = counts.get(j["Действие"], 0) + 1
    if run.unsorted_files:
        counts["не разобран"] = len(run.unsorted_files)   # файлов, а не строк по артикулам
    ws.append(["Параметр", "Значение"])
    for k, v in [("Запуск", run.stamp.strftime("%Y-%m-%d %H:%M:%S")),
                 ("Режим", ("пробный (--dry-run), ничего не изменено" if dry else
                            "боевой" if cfg.combat else "тест (Входящие не менялись)")),
                 ("Входящие", str(cfg.inbox)), ("Назначение", str(cfg.dest)),
                 ("Создано новых номенклатур", len(run.created))] + sorted(counts.items()):
        ws.append([k, v])
    for name, a, b in bank.dups:
        ws.append(["ВНИМАНИЕ: артикул в двух категориях", f"{name}: {a} и {b}"])
    style_sheet(ws, [40, 100])

    def table(title, rows, cols, widths):
        s = wb.create_sheet(title)
        s.append(cols)
        for r in rows:
            s.append([r.get(c, "") for c in cols])
        style_sheet(s, widths)

    table("Журнал", run.journal, ["Время", "Файл во «Входящих»", "Артикул", "Куда", "Действие", "Примечание"],
          [20, 60, 14, 90, 14, 60])
    table("Не разобрано", run.unsorted, ["Файл во «Входящих»", "Артикул", "Причина"], [70, 14, 90])
    if cfg.completeness:
        table("Полнота", completeness(cfg, bank),
              ["Категория", "Артикул", "Паспорт PDF", "Фото T1", "Фото товара, шт", "Сертификаты, шт",
               "Описание, шт", "Не хватает"], [40, 16, 12, 10, 14, 14, 12, 40])
    path = cfg.report
    try:
        wb.save(path)
    except PermissionError:
        path = path.with_name(f"{path.stem}_{run.stamp.strftime('%Y%m%d_%H%M%S')}{path.suffix}")
        wb.save(path)
        log(f"Отчёт открыт в Excel - сохранён под другим именем")
    return path


def append_journal(cfg, run):
    new = not cfg.journal.exists()
    cols = ["Время", "Файл во «Входящих»", "Артикул", "Куда", "Действие", "Примечание"]
    with open(cfg.journal, "a", encoding="utf-8-sig" if new else "utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols, delimiter=";")
        if new:
            w.writeheader()
        w.writerows(run.journal)


# --------------------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--book", default=BOOK_DEFAULT, help=f"книга правил (по умолчанию ./{BOOK_DEFAULT})")
    ap.add_argument("--init", action="store_true", help="создать книгу правил с настройками по умолчанию")
    ap.add_argument("--force", action="store_true", help="с --init: перезаписать существующую книгу")
    ap.add_argument("--dry-run", action="store_true", help="только показать, что будет сделано")
    args = ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    book = Path(args.book).resolve()
    if args.init:
        if book.exists() and not args.force:
            log(f"Книга уже есть: {book} (перезаписать: --init --force)")
            return 1
        create_book(book, book.parent / "structure.csv")
        log(f"Создана книга правил: {book}")
        return 0
    if not book.exists():
        log(f"Нет книги правил {book}. Создайте её: python inbox_sorter.py --init")
        return 1

    try:
        cfg = load_config(book)
    except (ValueError, OSError) as e:
        log(f"Ошибка в книге правил: {e}")
        return 1
    dry = args.dry_run
    log(f"Входящие:   {cfg.inbox}")
    log(f"Назначение: {cfg.dest}")
    log("Режим:      " + ("пробный (--dry-run)" if dry else "боевой" if cfg.combat else "тест"))
    if not os.path.isdir(ps.long_path(cfg.inbox)):
        if dry:
            log("Папки «Входящие» нет - разбирать нечего")
            return 0
        os.makedirs(ps.long_path(cfg.inbox), exist_ok=True)
    if not os.path.isdir(ps.long_path(cfg.dest)):
        if str(cfg.dest).startswith("\\\\"):
            log("Назначение недоступно. Проверьте, что NAS доступен и папка открывается в Проводнике.")
            return 1
        if not dry:
            os.makedirs(ps.long_path(cfg.dest), exist_ok=True)

    bank = scan_bank(cfg)
    log(f"В фотобанке номенклатур: {len(bank.items)}")
    for name, a, b in bank.dups:
        log(f"ВНИМАНИЕ: артикул {name} есть в двух местах: {a} и {b} - используется первое")
    resolver = Resolver(cfg, bank)
    run = Run(cfg, dry)
    errors = 0
    for rel in inbox_files(cfg):
        src = Path(cfg.inbox, *rel)
        src_rel = "/".join(rel)
        try:
            ok = process_file(run, resolver, src, rel)
        except OSError as e:
            errors += 1
            run.not_sorted(src_rel, f"ошибка: {e}")
            ok = None
        if dry or not cfg.combat:
            continue
        try:
            if ok:
                os.remove(ps.long_path(src))
            elif ok is False:
                target = unique_path(cfg.inbox / cfg.unsorted / Path(*rel))
                os.makedirs(ps.long_path(target.parent), exist_ok=True)
                shutil.move(ps.long_path(src), ps.long_path(target))
        except OSError as e:
            errors += 1
            log(f"  ОШИБКА при очистке «Входящих» {src_rel}: {e}")
    if cfg.combat and not dry:
        prune_empty_dirs(cfg.inbox, cfg.unsorted)

    report = write_report(cfg, run, bank, dry)
    if not dry:
        append_journal(cfg, run)
    done = sum(1 for j in run.journal if j["Действие"] in ("добавлен", "заменён"))
    log(f"Готово. Разложено файлов: {done}, уже было: "
        f"{sum(1 for j in run.journal if j['Действие'] == 'уже есть')}, "
        f"не разобрано: {len(run.unsorted_files)}, новых номенклатур: {len(run.created)}")
    log(f"Отчёт: {report}")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
