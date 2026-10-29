#!/usr/bin/env python3
"""
Раскладывает файлы из папки files по папкам номенклатурных единиц согласно structure.csv.

Для каждой строки structure.csv:
  <артикул>/                              <- "Имя корневой папки папки"
      <файлы для корневой папки>_<артикул>.ext
      <подпапка 1>/<файлы для подпапки 1>_<артикул>.ext
      <подпапка 2>/<файлы для подпапки 2>_<артикул>.ext
      <подпапка 3>/<имя файла>_<артикул>_3D.ext

Код модели перед артикулом (FO-R_, FO-DM_, FO-D_, FO-RM_ и т.п.) из имени
убирается: "Паспорт FO-R.pdf" -> "Паспорт 02-0331.pdf",
"FO-R.png" -> "02-0331_3D.png". Файлы, разложенные раньше со старыми именами,
переименовываются при каждом проходе.

Все папки артикулов создаются в отдельной папке "Номенклатура" рядом со
скриптом (можно задать другую через --dest).

Скрипт работает постоянно: каждые N секунд пересканирует папку files и
structure.csv и докопирует новые или изменившиеся файлы. Остановка - Ctrl+C.
"""

import argparse
import csv
import re
import shutil
import sys
import time
import unicodedata
from pathlib import Path

ROOT_COL = "Имя корневой папки папки"
SUBFOLDER_COLS = ["Имя для подпапки 1", "Имя для подпапки 2", "Имя для подпапки 3"]
ROOT_FILE_COLS = [f"Файл {i} для корневой папки" for i in range(1, 6)]
SUB1_FILE_COLS = ["Файл для подпапки 1", "Файл 2 для подпапки 1", "Файл 3 для подпапки 1"]
SUB2_FILE_COLS = ["Файл для подпапки 2", "Файл 2 для подпапки 2"]
SUB3_FILE_COLS = ["Файл для подпапки 3"]

# Файл считается дописанным, если не менялся столько секунд.
STABLE_SECONDS = 2.0


def log(msg):
    print(time.strftime("[%H:%M:%S] ") + msg, flush=True)


def normalize(name):
    return unicodedata.normalize("NFC", name).strip().casefold()


def read_structure(csv_path):
    raw = csv_path.read_bytes()
    for enc in ("utf-8-sig", "cp1251"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError(f"Не удалось определить кодировку {csv_path}")

    sample = text.split("\n", 1)[0]
    delimiter = ";" if sample.count(";") >= sample.count(",") else ","
    reader = csv.DictReader(text.splitlines(), delimiter=delimiter)
    reader.fieldnames = [(f or "").strip() for f in reader.fieldnames or []]
    if ROOT_COL not in reader.fieldnames:
        raise ValueError(f"В {csv_path.name} нет колонки '{ROOT_COL}'")
    rows = []
    for row in reader:
        row = {k: (v or "").strip() for k, v in row.items() if k}
        if row.get(ROOT_COL):
            rows.append(row)
    return rows


def strip_model_code(name, article):
    """Убирает код модели (латиница/цифры/дефисы, напр. FO-R) прямо перед _<артикул>."""
    pattern = r"(^|[ _])[A-Z][A-Za-z0-9]*(?:-[A-Za-z0-9]+)*_(?=" + re.escape(article) + ")"
    return re.sub(pattern, r"\1", name)


def build_targets(row):
    """Возвращает список (имя исходного файла, относительная папка, новое имя)."""
    article = row[ROOT_COL]
    groups = [
        ("", ROOT_FILE_COLS, ""),
        (row.get(SUBFOLDER_COLS[0], ""), SUB1_FILE_COLS, ""),
        (row.get(SUBFOLDER_COLS[1], ""), SUB2_FILE_COLS, ""),
        (row.get(SUBFOLDER_COLS[2], ""), SUB3_FILE_COLS, "_3D"),
    ]
    targets = []
    for subfolder, cols, suffix in groups:
        for col in cols:
            src_name = row.get(col, "")
            if not src_name:
                continue
            p = Path(src_name)
            new_name = strip_model_code(f"{p.stem}_{article}{suffix}{p.suffix}", article)
            targets.append((src_name, subfolder, new_name))
    return targets


def index_sources(files_dir):
    """Словарь: нормализованное имя -> путь; плюс по имени без расширения."""
    by_name, by_stem = {}, {}
    if not files_dir.is_dir():
        return by_name, by_stem
    for p in files_dir.iterdir():
        if p.is_file():
            by_name[normalize(p.name)] = p
            by_stem.setdefault(normalize(p.stem), p)
    return by_name, by_stem


def find_source(name, by_name, by_stem):
    key = normalize(name)
    return by_name.get(key) or by_stem.get(key) or by_stem.get(normalize(Path(name).stem))


def needs_copy(src, dst):
    if not dst.exists():
        return True
    s, d = src.stat(), dst.stat()
    return s.st_size != d.st_size or s.st_mtime > d.st_mtime + 1


def clean_existing(root, article, dest):
    """Переименовывает уже разложенные файлы, в именах которых остался код модели."""
    renamed = 0
    if not root.is_dir():
        return renamed
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        clean = strip_model_code(path.name, article)
        if clean == path.name:
            continue
        target = path.with_name(clean)
        if target.exists():
            path.unlink()  # чистая версия уже есть - старая копия не нужна
        else:
            path.rename(target)
        renamed += 1
        log(f"Переименован: {path.relative_to(dest)} -> {clean}")
    return renamed


def process(csv_path, files_dir, dest, reported_missing):
    rows = read_structure(csv_path)
    by_name, by_stem = index_sources(files_dir)
    now = time.time()
    copied = 0

    for row in rows:
        article = row[ROOT_COL]
        root = dest / article
        clean_existing(root, article, dest)
        for src_name, subfolder, new_name in build_targets(row):
            src = find_source(src_name, by_name, by_stem)
            if src is None:
                key = (article, src_name)
                if key not in reported_missing:
                    reported_missing.add(key)
                    log(f"Нет файла в files: '{src_name}' (артикул {article})")
                continue
            reported_missing.discard((article, src_name))
            if now - src.stat().st_mtime < STABLE_SECONDS:
                continue  # файл ещё копируется/дописывается, возьмём на следующем проходе
            target_dir = root / subfolder if subfolder else root
            dst = target_dir / new_name
            if not needs_copy(src, dst):
                continue
            target_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
            copied += 1
            log(f"{src.name} -> {dst.relative_to(dest)}")

        # Создаём пустые подпапки, даже если файлов для них пока нет.
        for col in SUBFOLDER_COLS:
            if row.get(col):
                (root / row[col]).mkdir(parents=True, exist_ok=True)
    return copied


def main():
    base = Path.cwd()
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--files", type=Path, default=base / "files", help="папка с исходными файлами (по умолчанию ./files)")
    ap.add_argument("--csv", type=Path, default=base / "structure.csv", help="файл связей (по умолчанию ./structure.csv)")
    ap.add_argument("--dest", type=Path, default=base / "Номенклатура", help="куда раскладывать (по умолчанию ./Номенклатура)")
    ap.add_argument("--interval", type=float, default=5.0, help="пауза между сканированиями, сек (по умолчанию 5)")
    ap.add_argument("--once", action="store_true", help="выполнить один проход и выйти")
    args = ap.parse_args()

    files_dir, csv_path, dest = args.files.resolve(), args.csv.resolve(), args.dest.resolve()
    files_dir.mkdir(parents=True, exist_ok=True)
    dest.mkdir(parents=True, exist_ok=True)
    log(f"Исходные файлы: {files_dir}")
    log(f"Структура:      {csv_path}")
    log(f"Назначение:     {dest}")

    reported_missing = set()
    last_error = None
    try:
        while True:
            try:
                copied = process(csv_path, files_dir, dest, reported_missing)
                if copied:
                    log(f"Скопировано файлов: {copied}")
                last_error = None
            except (OSError, ValueError, csv.Error) as e:
                if str(e) != last_error:
                    log(f"Ошибка: {e}")
                    last_error = str(e)
            if args.once:
                break
            time.sleep(args.interval)
    except KeyboardInterrupt:
        log("Остановлено.")


if __name__ == "__main__":
    sys.exit(main())
