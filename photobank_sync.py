#!/usr/bin/env python3
"""
Копирует фотобанк с Synology (SMB) на локальный диск, приводя папку каждой
номенклатуры к единой структуре:

    <категория>/<артикул>/
        Паспорт <артикул>.pdf
        Описание/
        Сертификаты/
        Служебные материалы/
            Исходники/
            Протокол испытаний/
            Старые фото/
            Упаковка/
        Фото товара/

Источник: <src>/<категория>/<артикул>/... (по умолчанию
\\\\192.168.13.200\\дм\\ОПМ\\БАЗА ПО ПРОДУКТАМ - ФОТОБАНК).

Копируются только папки категорий из CATEGORIES (1-10); остальное
содержимое корня фотобанка пропускается.

Все папки структуры создаются всегда; если каких-то файлов нет, остаются
пустые папки. Файлы, уже лежащие в папках с правильными именами, сохраняют
своё место; файлы из корня или из нестандартных папок раскладываются по
правилам RULES ниже. Нераспознанные файлы из корня номенклатуры остаются
в корне, из нестандартных папок - попадают в "Служебные материалы" с
сохранением исходного пути. Любая папка "Старые фото" переносится целиком
в "Служебные материалы/Старые фото" без изменения содержимого. Служебный мусор Synology/Windows/macOS
(._DAV, @eaDir, Thumbs.db, .DS_Store ...) не копируется.

Повторный запуск докопирует только новые и изменённые файлы. Исходные
файлы на NAS не изменяются.
"""

import argparse
import os
import re
import shutil
import sys
import time
import unicodedata
from pathlib import Path

DEFAULT_SRC = r"\\192.168.13.200\дм\ОПМ\БАЗА ПО ПРОДУКТАМ - ФОТОБАНК"
DEFAULT_DEST = "Фотобанк"

# Какие папки фотобанка копировать: только категории номенклатуры. Папка
# берётся, если её имя начинается с номера из списка и пробела ("1 LAN CCA ...",
# "10 Телекоммуникационные ..."); всё остальное в корне фотобанка пропускается.
# Сопоставление по номеру, а не по полному имени, чтобы мелкие правки названий
# категорий на NAS не ломали копирование.
CATEGORIES = {
    1: "LAN CCA кабель стандартные упаковки",
    2: "LAN-кабель стандартные упаковки",
    3: "Арматура и монтажные материалы",
    4: "Инструменты",
    5: "Коаксиальный кабель в стандартной упаковке",
    6: "Коннекторы, разъемы, переходники",
    7: "Медные компоненты",
    8: "Оптические кабели стандартные упаковки",
    9: "Оптические компоненты",
    10: "Телекоммуникационные шкафы, стойки",
}
CATEGORY_RE = re.compile(r"^\s*(\d+)\s+\S")

SERVICE = "Служебные материалы"
# Структура папки номенклатуры (по folder_example.zip).
TEMPLATE = [
    "Описание",
    "Сертификаты",
    SERVICE,
    f"{SERVICE}/Исходники",
    f"{SERVICE}/Протокол испытаний",
    f"{SERVICE}/Старые фото",
    f"{SERVICE}/Упаковка",
    "Фото товара",
]
# Эти папки не создаются пустыми - только если в них есть что положить.
OPTIONAL_DIRS = {f"{SERVICE}/Исходники", f"{SERVICE}/Старые фото"}

# Документы для папки "Сертификаты": декларации, сертификаты (в т.ч. пожарные),
# отказные письма, уведомления об отказе, отрицательные решения, выписки из
# реестра российской продукции (ТОРП), свидетельства.
CERT_KEYWORDS = ("сертификат", "декларац", "отказ", "отрицательн", "выписк", "реестр", "свидетельств")

JUNK_DIRS = {"._dav", "@eadir", "#recycle", "#snapshot", ".spotlight-v100", ".trashes", ".fseventsd"}
JUNK_FILES = {"thumbs.db", ".ds_store", "desktop.ini", "ehthumbs.db"}

RAW_EXT = {".cr2", ".cr3", ".nef", ".arw", ".orf", ".raf", ".rw2", ".dng",
           ".psd", ".psb", ".tif", ".tiff", ".ai", ".eps", ".indd"}
IMAGE_EXT = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"}
PHOTO_EXT = IMAGE_EXT | {".tif", ".tiff"}

# Папка "Старые фото", где бы она ни лежала в папке номенклатуры, переносится
# в OLD_PHOTOS_TARGET целиком: те же подпапки и имена файлов, без раскладки
# по правилам.
OLD_PHOTOS_TARGET = f"{SERVICE}/Старые фото"
OLD_PHOTOS_DIR = "старые фото"


def is_product_photo(name, article):
    """<артикул>_T1, _Y2, _3D, _G4 ... (имя и артикул уже в нижнем регистре)."""
    return re.fullmatch(re.escape(article) + r"_(?:[a-zа-яё]+\d*|3d\d*)\.[^.]+", name) is not None
DOC_EXT = {".doc", ".docx", ".rtf", ".odt", ".txt"}

# Правила для файлов, лежащих не в стандартной папке. Проверяются по порядку,
# первое совпавшее определяет папку ("" - корень папки номенклатуры).
# Аргументы: имя файла в нижнем регистре, расширение, артикул в нижнем регистре.
RULES = [
    # Готовые фото товара, названные по артикулу: 05-0900_T1, _T6, _Y1, _Y2, _3D ...
    # Стоит первым, чтобы, например, 05-0900_3D.tif не ушёл в "Исходники".
    (lambda n, e, a: e in PHOTO_EXT and is_product_photo(n, a), "Фото товара"),
    (lambda n, e, a: n.startswith("паспорт") and e == ".pdf", ""),
    (lambda n, e, a: n.startswith("паспорт"), SERVICE),
    (lambda n, e, a: any(k in n for k in CERT_KEYWORDS), "Сертификаты"),
    (lambda n, e, a: "протокол" in n, f"{SERVICE}/Протокол испытаний"),
    (lambda n, e, a: "описани" in n, "Описание"),
    (lambda n, e, a: any(k in n for k in ("этикет", "упаков", "коробк")), f"{SERVICE}/Упаковка"),
    (lambda n, e, a: "сплеш" in n or "splash" in n, SERVICE),
    (lambda n, e, a: e in RAW_EXT, f"{SERVICE}/Исходники"),
]


def log(msg):
    print(time.strftime("[%H:%M:%S] ") + msg, flush=True)


def norm(name):
    return unicodedata.normalize("NFC", name).strip().casefold()


# Нормализованный путь шаблонной папки -> каноническое написание.
CANON = {tuple(norm(p) for p in t.split("/")): t for t in TEMPLATE}


def long_path(p):
    """Префикс \\\\?\\ снимает ограничение Windows в 260 символов."""
    s = str(p)
    if os.name != "nt" or s.startswith("\\\\?\\"):
        return s
    s = os.path.abspath(s)
    if s.startswith("\\\\"):
        return "\\\\?\\UNC\\" + s[2:]
    return "\\\\?\\" + s


def is_junk(name, is_dir):
    n = norm(name)
    if n.startswith("._") or n.startswith("~$"):
        return True
    return n in JUNK_DIRS if is_dir else n in JUNK_FILES


def list_dir(path):
    """Возвращает (папки, файлы) без служебного мусора, отсортированные по имени."""
    dirs, files = [], []
    with os.scandir(long_path(path)) as it:
        for e in it:
            try:
                d = e.is_dir()
            except OSError:
                continue
            if is_junk(e.name, d):
                continue
            (dirs if d else files).append(e.name)
    return sorted(dirs), sorted(files)


def walk_files(root, dirs_out=None):
    """Все файлы внутри root (относительные пути как кортежи частей), без мусора.
    Если передан dirs_out, в него добавляются пути всех найденных папок."""
    stack = [()]
    while stack:
        rel = stack.pop()
        dirs, files = list_dir(Path(root, *rel))
        for f in files:
            yield rel + (f,)
        for d in reversed(dirs):
            stack.append(rel + (d,))
            if dirs_out is not None:
                dirs_out.append(rel + (d,))


def classify(rel_parts, article):
    """Определяет относительный путь (в папке номенклатуры), куда положить файл."""
    name = rel_parts[-1]
    folders = rel_parts[:-1]
    nfolders = tuple(norm(f) for f in folders)

    # "Старые фото" (на любом уровне) - целиком в утверждённое место,
    # содержимое и вложенные папки не трогаем.
    if OLD_PHOTOS_DIR in nfolders:
        k = nfolders.index(OLD_PHOTOS_DIR)
        return Path(OLD_PHOTOS_TARGET, *folders[k + 1:], name)

    # Файл уже лежит в стандартной папке - берём самый длинный совпавший префикс.
    for k in range(len(nfolders), 0, -1):
        canon = CANON.get(nfolders[:k])
        if canon:
            return Path(canon, *folders[k:], name)

    n, e, a = norm(name), Path(name).suffix.lower(), norm(article)
    for rule, target in RULES:
        if rule(n, e, a):
            return Path(target, name) if target else Path(name)
    # Не распознали: из корня номенклатуры - оставляем в корне как есть,
    # из нестандартной папки - в служебные материалы, сохраняя исходные подпапки.
    if not folders:
        return Path(name)
    return Path(SERVICE, *folders, name)


def is_category(name):
    m = CATEGORY_RE.match(name)
    return bool(m) and int(m.group(1)) in CATEGORIES


def needs_copy(src, dst):
    try:
        d = os.stat(long_path(dst))
    except FileNotFoundError:
        return True
    s = os.stat(long_path(src))
    return s.st_size != d.st_size or s.st_mtime > d.st_mtime + 2


def unique_target(rel, used):
    """Если два исходных файла попадают в одно место - добавляем (2), (3)..."""
    key = norm(str(rel))
    if key not in used:
        used.add(key)
        return rel
    i = 2
    while True:
        cand = rel.with_name(f"{rel.stem} ({i}){rel.suffix}")
        k = norm(str(cand))
        if k not in used:
            used.add(k)
            return cand
        i += 1


def sync_item(src_item, dst_item, article, stats, dry_run):
    for t in TEMPLATE:
        if t not in OPTIONAL_DIRS and not dry_run:
            os.makedirs(long_path(dst_item / t), exist_ok=True)

    # Необязательные папки создаём, если такая папка есть в источнике (даже
    # пустая); если в неё попадут файлы, она создастся при копировании.
    src_dirs = []
    files = list(walk_files(src_item, src_dirs))
    optional = {norm(o): o for o in OPTIONAL_DIRS}
    for d in src_dirs:
        target = optional.get(norm(classify(d + ("_",), article).parent.as_posix()))
        if target and norm(d[-1]) == norm(Path(target).name) and not dry_run:
            os.makedirs(long_path(dst_item / target), exist_ok=True)

    used = set()
    for rel_parts in files:
        src = Path(src_item, *rel_parts)
        rel = unique_target(classify(rel_parts, article), used)
        dst = dst_item / rel
        moved = Path(*rel_parts) != rel
        try:
            if not needs_copy(src, dst):
                stats["same"] += 1
                continue
            note = f"  (из {Path(*rel_parts)})" if moved else ""
            log(f"  {article}/{rel.as_posix()}{note}")
            if not dry_run:
                os.makedirs(long_path(dst.parent), exist_ok=True)
                shutil.copy2(long_path(src), long_path(dst))
            stats["copied"] += 1
        except OSError as ex:
            stats["errors"] += 1
            log(f"  ОШИБКА {src}: {ex}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--src", default=DEFAULT_SRC, help="корень фотобанка на NAS")
    ap.add_argument("--dest", default=DEFAULT_DEST, help=f"куда копировать (по умолчанию ./{DEFAULT_DEST})")
    ap.add_argument("--only", nargs="*", default=None,
                    help="обработать только эти категории/артикулы (подстрока имени)")
    ap.add_argument("--dry-run", action="store_true", help="только показать, что будет сделано")
    args = ap.parse_args()

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    src_root, dest_root = Path(args.src), Path(args.dest).resolve()
    if not os.path.isdir(long_path(src_root)):
        log(f"Источник недоступен: {src_root}")
        log("Проверьте, что NAS доступен и папка открывается в Проводнике.")
        return 1
    log(f"Источник:   {src_root}")
    log(f"Назначение: {dest_root}{'  (пробный запуск)' if args.dry_run else ''}")

    only = [norm(o) for o in args.only] if args.only else None
    stats = {"items": 0, "copied": 0, "same": 0, "errors": 0}
    all_dirs, _ = list_dir(src_root)
    categories = sorted((d for d in all_dirs if is_category(d)),
                        key=lambda d: int(CATEGORY_RE.match(d).group(1)))
    skipped = [d for d in all_dirs if d not in categories]
    log(f"Категорий к копированию: {len(categories)}, прочих папок пропущено: {len(skipped)}")
    found = {int(CATEGORY_RE.match(d).group(1)) for d in categories}
    for num in sorted(set(CATEGORIES) - found):
        log(f"  ВНИМАНИЕ: не найдена категория {num} {CATEGORIES[num]}")
    for cat in categories:
        try:
            items, loose = list_dir(src_root / cat)
        except OSError as ex:
            log(f"ОШИБКА чтения {cat}: {ex}")
            stats["errors"] += 1
            continue
        if only and not any(o in norm(cat) for o in only):
            items = [i for i in items if any(o in norm(i) for o in only)]
            if not items:
                continue
        log(f"[{cat}] номенклатур: {len(items)}")
        if loose:
            log(f"  файлы в корне категории пропущены: {len(loose)}")
        for item in items:
            stats["items"] += 1
            try:
                sync_item(src_root / cat / item, dest_root / cat / item, item, stats, args.dry_run)
            except OSError as ex:
                stats["errors"] += 1
                log(f"  ОШИБКА {cat}/{item}: {ex}")

    log(f"Готово. Номенклатур: {stats['items']}, скопировано файлов: {stats['copied']}, "
        f"без изменений: {stats['same']}, ошибок: {stats['errors']}")
    return 1 if stats["errors"] else 0


if __name__ == "__main__":
    sys.exit(main())
