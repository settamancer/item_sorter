#!/usr/bin/env python3
"""
Раскладывает файлы из папки files по папкам номенклатурных единиц согласно structure.csv.

Для каждой строки structure.csv:
  <артикул>/                              <- "Имя корневой папки папки"
      <файлы для корневой папки>_<артикул>.ext
      <подпапка 1>/<файлы для подпапки 1>_<артикул>.ext
      <подпапка 2>/<файлы для подпапки 2>_<артикул>.ext
      <подпапка 3>/<имя файла>_<артикул>_3D.ext

Папка артикула ищется в целевом каталоге (на 1-2 уровня вглубь, например
"2 LAN-кабель стандартные упаковки/02-0331"). Если её нет, она создаётся
в папке категории, номер которой совпадает с префиксом артикула
("02-0331" -> "2 ..."), а если такой категории нет - прямо в целевом каталоге.

Скрипт работает постоянно: каждые N секунд пересканирует папку files и
structure.csv и докопирует новые или изменившиеся файлы. Остановка - Ctrl+C.
"""

import argparse
import csv
import os
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
            new_name = f"{p.stem}_{article}{suffix}{p.suffix}"
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


def category_number(name):
    head = name.split(" ", 1)[0].split("-", 1)[0]
    return int(head) if head.isdigit() else None


class FolderLocator:
    """Ищет/выбирает папку артикула в целевом каталоге."""

    def __init__(self, dest, skip):
        self.dest = dest
        self.skip = {s.resolve() for s in skip}
        self.refresh()

    def refresh(self):
        self.existing = {}
        self.categories = {}
        for d in self._subdirs(self.dest):
            self.existing.setdefault(d.name, d)
            num = category_number(d.name)
            if num is not None and " " in d.name:
                self.categories.setdefault(num, d)
            for sub in self._subdirs(d):
                self.existing.setdefault(sub.name, sub)

    def _subdirs(self, path):
        try:
            return sorted(
                p for p in path.iterdir()
                if p.is_dir() and not p.name.startswith(".") and p.resolve() not in self.skip
            )
        except OSError:
            return []

    def locate(self, article):
        if article in self.existing:
            return self.existing[article]
        num = category_number(article)
        parent = self.categories.get(num, self.dest)
        folder = parent / article
        self.existing[article] = folder
        return folder


def needs_copy(src, dst):
    if not dst.exists():
        return True
    s, d = src.stat(), dst.stat()
    return s.st_size != d.st_size or s.st_mtime > d.st_mtime + 1


def process(csv_path, files_dir, dest, reported_missing):
    rows = read_structure(csv_path)
    by_name, by_stem = index_sources(files_dir)
    locator = FolderLocator(dest, skip=[files_dir])
    now = time.time()
    copied = 0

    for row in rows:
        article = row[ROOT_COL]
        root = locator.locate(article)
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
    ap.add_argument("--dest", type=Path, default=base, help="куда раскладывать (по умолчанию текущая папка)")
    ap.add_argument("--interval", type=float, default=5.0, help="пауза между сканированиями, сек (по умолчанию 5)")
    ap.add_argument("--once", action="store_true", help="выполнить один проход и выйти")
    args = ap.parse_args()

    files_dir, csv_path, dest = args.files.resolve(), args.csv.resolve(), args.dest.resolve()
    files_dir.mkdir(parents=True, exist_ok=True)
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
