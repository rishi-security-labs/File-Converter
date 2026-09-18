#!/usr/bin/env python3
"""
RISHI DATA CONVERTER
JSON ↔ CSV ↔ XLSX ↔ SQLite ↔ SQL
Colorful Termux / Android terminal application.
"""

from __future__ import annotations

import csv
import json
import os
import re
import sqlite3
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Iterator


APP_TITLE = "RISHI DATA CONVERTER"
BOX_W = 54
PAGE_SIZE = 14
PROGRESS_EVERY = 2000
SETTINGS_PATH = Path.home() / ".rishi_data_converter" / "settings.json"
BATCH_INSERT = 400

FORMATS = ("json", "csv", "xlsx", "sqlite", "sql")
FORMAT_LABEL = {
    "json": "JSON",
    "csv": "CSV",
    "xlsx": "EXCEL (XLSX)",
    "sqlite": "SQLITE",
    "sql": "SQL",
}
FORMAT_EXT = {
    "json": ".json",
    "csv": ".csv",
    "xlsx": ".xlsx",
    "sqlite": ".db",
    "sql": ".sql",
}
FORMAT_EXTS = {
    "json": {".json"},
    "csv": {".csv"},
    "xlsx": {".xlsx"},
    "sqlite": {".db", ".sqlite", ".sqlite3"},
    "sql": {".sql"},
}
ALL_EXTS = set().union(*FORMAT_EXTS.values())

SQLITE_FILE_EXTS = FORMAT_EXTS["sqlite"]


class C:
    RESET = "\033[0m"
    BOLD = "\033[1m"
    DIM = "\033[2m"
    RED = "\033[31m"
    GREEN = "\033[32m"
    YELLOW = "\033[33m"
    BLUE = "\033[34m"
    MAGENTA = "\033[35m"
    CYAN = "\033[36m"
    WHITE = "\033[37m"
    BRED = "\033[91m"
    BGREEN = "\033[92m"
    BYELLOW = "\033[93m"
    BBLUE = "\033[94m"
    BMAGENTA = "\033[95m"
    BCYAN = "\033[96m"
    BWHITE = "\033[97m"


ANSI_RE = re.compile(r"\033\[[0-9;]*m")
_OPENPYXL = None
_COLORS = True


def apply_color_mode(enabled: bool) -> None:
    global _COLORS
    _COLORS = enabled
    codes = {
        "RESET": "\033[0m",
        "BOLD": "\033[1m",
        "DIM": "\033[2m",
        "RED": "\033[31m",
        "GREEN": "\033[32m",
        "YELLOW": "\033[33m",
        "BLUE": "\033[34m",
        "MAGENTA": "\033[35m",
        "CYAN": "\033[36m",
        "WHITE": "\033[37m",
        "BRED": "\033[91m",
        "BGREEN": "\033[92m",
        "BYELLOW": "\033[93m",
        "BBLUE": "\033[94m",
        "BMAGENTA": "\033[95m",
        "BCYAN": "\033[96m",
        "BWHITE": "\033[97m",
    }
    for name, code in codes.items():
        setattr(C, name, code if enabled else "")


def vis_len(text: str) -> int:
    return len(ANSI_RE.sub("", text))


def clip(text: str, width: int) -> str:
    raw = ANSI_RE.sub("", text)
    if len(raw) <= width:
        return text + (" " * (width - len(raw)))
    return raw[: max(0, width - 1)] + "…"


def center(text: str, width: int = BOX_W) -> str:
    n = vis_len(text)
    if n >= width:
        return text
    left = (width - n) // 2
    return (" " * left) + text + (" " * (width - n - left))


def clear() -> None:
    os.system("clear")


def pause(msg: str = "Press Enter to continue...") -> None:
    try:
        input(f"\n{C.DIM}{msg}{C.RESET}")
    except EOFError:
        pass


def ask(prompt: str) -> str:
    try:
        return input(prompt).strip()
    except EOFError:
        return ""
    except KeyboardInterrupt:
        print()
        raise


def box_top() -> None:
    print(f"{C.CYAN}╔{'═' * BOX_W}╗{C.RESET}")


def box_mid() -> None:
    print(f"{C.CYAN}╠{'═' * BOX_W}╣{C.RESET}")


def box_bot() -> None:
    print(f"{C.CYAN}╚{'═' * BOX_W}╝{C.RESET}")


def box_row(text: str, color: str = "") -> None:
    print(f"{C.CYAN}║{C.RESET}{color}{clip(text, BOX_W)}{C.RESET}{C.CYAN}║{C.RESET}")


def box_center(text: str, color: str = "") -> None:
    box_row(center(text), color)


def error_box(message: str) -> None:
    print()
    box_top()
    box_center("❌ ERROR", C.BOLD + C.BRED)
    box_mid()
    for line in str(message).splitlines() or ["Unknown error"]:
        box_row(" " + line, C.WHITE)
    box_bot()


def success_box(rows: int, cols: int, files: list[Path], folder: Path) -> None:
    print()
    box_top()
    box_center("✅ COMPLETE", C.BOLD + C.BGREEN)
    box_mid()
    box_row(f" 📊 Rows       : {rows:,}", C.WHITE)
    box_row(f" 📋 Columns    : {cols}", C.WHITE)
    box_row(f" 📦 Files      : {len(files)}", C.WHITE)
    if len(files) == 1:
        box_row(f" 📄 Output     : {files[0].name}", C.WHITE)
    box_row(f" 📁 Folder     : {folder}", C.WHITE)
    if len(files) > 1:
        box_mid()
        shown = files[:8]
        for p in shown:
            try:
                sz = format_size(p.stat().st_size)
            except OSError:
                sz = "?"
            box_row(f" → {p.name}  {sz}", C.BGREEN)
        if len(files) > 8:
            box_row(f" → ... +{len(files) - 8} more", C.DIM)
    box_bot()


def banner(subtitle: str = "") -> None:
    print(f"{C.BOLD}{C.BMAGENTA}", end="")
    box_top()
    box_row("")
    box_center("R I S H I", C.BOLD + C.BCYAN)
    box_row("")
    box_center("DATA CONVERTER TOOL", C.BOLD + C.BYELLOW)
    box_row("")
    box_center("JSON ↔ CSV ↔ XLSX ↔ SQLite ↔ SQL", C.BGREEN)
    box_row("")
    box_bot()
    print(C.RESET, end="")
    if subtitle:
        print(f"{C.MAGENTA}{subtitle}{C.RESET}")


def format_size(num_bytes: int | float) -> str:
    n = float(num_bytes)
    if n >= 1024 * 1024 * 1024:
        return f"{n / (1024 * 1024 * 1024):.2f} GB"
    if n >= 1024 * 1024:
        return f"{n / (1024 * 1024):.2f} MB"
    if n >= 1024:
        return f"{n / 1024:.2f} KB"
    return f"{int(n)} B"


def parse_size(text: str) -> int:
    raw = text.strip().upper().replace(" ", "")
    if not raw:
        raise ValueError("Size empty hai.")
    m = re.fullmatch(r"([0-9]*\.?[0-9]+)([KMG]?B)?", raw)
    if not m:
        raise ValueError("Size samajh nahi aayi. Example: 100 KB  ya  1 MB")
    value = float(m.group(1))
    unit = m.group(2) or "MB"
    if value <= 0:
        raise ValueError("Size 0 se badi honi chahiye.")
    mult = {"B": 1, "KB": 1024, "MB": 1024 ** 2, "GB": 1024 ** 3}
    if unit not in mult:
        raise ValueError("Unit B / KB / MB / GB use karo.")
    size = int(value * mult[unit])
    if size < 512:
        raise ValueError("Size bahut chhoti hai. Kam se kam 512 B do.")
    return size


def default_settings() -> dict:
    return {
        "colors": True,
        "default_output": "",
        "default_split": "1 MB",
    }


def load_settings() -> dict:
    data = default_settings()
    try:
        if SETTINGS_PATH.exists():
            loaded = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                data.update({k: loaded[k] for k in data if k in loaded})
    except Exception:
        pass
    apply_color_mode(bool(data.get("colors", True)))
    return data


def save_settings(data: dict) -> None:
    SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
    SETTINGS_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")


SETTINGS = load_settings()


def storage_roots() -> list[Path]:
    candidates = [
        Path("/sdcard"),
        Path("/storage/emulated/0"),
        Path.home() / "storage" / "shared",
        Path.home() / "storage" / "downloads",
        Path("/sdcard/Download"),
        Path.home(),
        Path("/"),
    ]
    seen = set()
    out = []
    for p in candidates:
        try:
            rp = p.resolve() if p.exists() else p
        except OSError:
            rp = p
        key = str(rp)
        if key in seen:
            continue
        seen.add(key)
        if p.exists():
            out.append(p)
    return out or [Path.home()]


def home_path() -> Path:
    for p in storage_roots():
        if str(p) != "/" and os.access(p, os.R_OK):
            return p
    return Path.home()


def check_storage() -> None:
    readable = False
    for p in (Path("/sdcard"), Path("/storage/emulated/0"), Path.home() / "storage" / "shared"):
        if p.exists() and os.access(p, os.R_OK):
            readable = True
            break
    if not readable:
        error_box(
            "Android storage accessible nahi hai.\n"
            "Termux mein ye command chalao:\n"
            "termux-setup-storage"
        )


def safe_filename(name: str) -> str:
    name = name.strip().replace("\\", "_").replace("/", "_")
    name = re.sub(r"[^\w.\- +()]+", "_", name, flags=re.UNICODE).strip(" .")
    return name or "output"


def safe_ident(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]", "_", str(name or "").strip())
    if not cleaned:
        cleaned = "table1"
    if cleaned[0].isdigit():
        cleaned = "t_" + cleaned
    return cleaned[:60]


def quote_ident(name: str) -> str:
    return '"' + safe_ident(name).replace('"', "") + '"'


def is_empty_value(value) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def clean_row(obj: dict) -> dict:
    out = {}
    if not isinstance(obj, dict):
        return out
    for key, value in obj.items():
        if key is None:
            continue
        k = str(key).strip()
        if not k:
            continue
        if is_empty_value(value):
            continue
        if isinstance(value, str):
            value = value.strip()
            if value == "":
                continue
        out[k] = value
    return out


def normalize_cell(value):
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return value
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)
    text = str(value).strip()
    return text if text else None


def unique_columns(rows: Iterable[dict]) -> list[str]:
    cols: list[str] = []
    seen = set()
    for row in rows:
        for k in row.keys():
            if k not in seen:
                seen.add(k)
                cols.append(k)
    return cols


def alloc_path(folder: Path, stem: str, ext: str, index: int | None = None) -> Path:
    stem = safe_filename(stem)
    if not ext.startswith("."):
        ext = "." + ext
    n = 0 if index is None else index
    while True:
        name = f"{stem}{ext}" if n == 0 else f"{stem}{n}{ext}"
        path = folder / name
        if not path.exists():
            return path
        n += 1


def get_unique_folder(parent: Path, name: str) -> Path:
    safe = safe_filename(name)
    candidate = parent / safe
    if not candidate.exists():
        candidate.mkdir(parents=True, exist_ok=True)
        return candidate
    n = 1
    while True:
        candidate = parent / f"{safe}{n}"
        if not candidate.exists():
            candidate.mkdir(parents=True, exist_ok=True)
            return candidate
        n += 1


@dataclass
class Table:
    name: str
    columns: list[str] = field(default_factory=list)
    rows: list[dict] = field(default_factory=list)

    def finalize(self) -> None:
        if not self.columns:
            self.columns = unique_columns(self.rows)
        cleaned = []
        for row in self.rows:
            item = {}
            for k, v in row.items():
                nv = normalize_cell(v)
                if is_empty_value(nv):
                    continue
                item[str(k)] = nv
            if item:
                cleaned.append(item)
        self.rows = cleaned
        if not self.columns:
            self.columns = unique_columns(self.rows)

    @property
    def count(self) -> int:
        return len(self.rows)


def infer_sql_type(col: str, rows: list[dict]) -> str:
    seen_int = seen_real = seen_text = False
    checked = 0
    for row in rows:
        if col not in row:
            continue
        v = row[col]
        if is_empty_value(v):
            continue
        checked += 1
        if isinstance(v, bool):
            seen_text = True
        elif isinstance(v, int):
            seen_int = True
        elif isinstance(v, float):
            seen_real = True
        else:
            s = str(v).strip()
            if re.fullmatch(r"-?\d+", s):
                seen_int = True
            elif re.fullmatch(r"-?\d+\.\d+", s):
                seen_real = True
            else:
                seen_text = True
        if checked >= 80:
            break
    if seen_text:
        return "TEXT"
    if seen_real:
        return "REAL"
    if seen_int:
        return "INTEGER"
    return "TEXT"


def sql_literal(value) -> str:
    if value is None:
        return "NULL"
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, int) and not isinstance(value, bool):
        return str(value)
    if isinstance(value, float):
        return repr(value)
    return "'" + str(value).replace("'", "''") + "'"


def ensure_openpyxl():
    global _OPENPYXL
    if _OPENPYXL is not None:
        return _OPENPYXL
    try:
        import openpyxl  # type: ignore

        _OPENPYXL = openpyxl
        return _OPENPYXL
    except ImportError:
        pass
    print()
    box_top()
    box_center("⚠️  openpyxl REQUIRED", C.BOLD + C.BYELLOW)
    box_mid()
    box_row(" Excel support ke liye openpyxl chahiye.", C.WHITE)
    box_row(" Command: python -m pip install openpyxl", C.WHITE)
    box_mid()
    box_row("  1. Install now", C.BGREEN)
    box_row("  2. Cancel", C.BRED)
    box_bot()
    if ask(f"\n{C.YELLOW}Select: {C.RESET}") != "1":
        raise RuntimeError("Excel conversion cancel. openpyxl missing hai.")
    print(f"\n{C.YELLOW}🔄 Installing openpyxl...{C.RESET}\n")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
        import openpyxl  # type: ignore

        _OPENPYXL = openpyxl
        print(f"{C.BGREEN}✅ openpyxl install ho gaya.{C.RESET}")
        return _OPENPYXL
    except Exception as exc:
        raise RuntimeError(f"openpyxl install fail.\n{exc}") from exc


def list_dir(path: Path, exts: set[str] | None) -> tuple[list[Path], list[Path]]:
    folders: list[Path] = []
    files: list[Path] = []
    try:
        entries = list(path.iterdir())
    except PermissionError as exc:
        raise PermissionError(f"Permission denied: {path}") from exc
    except OSError as exc:
        raise OSError(f"Folder read nahi hua: {path}\n{exc}") from exc
    for ent in entries:
        try:
            name = ent.name
            if name.startswith("."):
                continue
            if ent.is_dir():
                folders.append(ent)
            elif ent.is_file() and exts is not None:
                if ent.suffix.lower() in exts:
                    files.append(ent)
        except OSError:
            continue
    folders.sort(key=lambda p: p.name.lower())
    files.sort(key=lambda p: p.name.lower())
    return folders, files


class NavCancel(Exception):
    pass


class NavHome(Exception):
    pass


class NavBack(Exception):
    pass


def handle_nav(raw: str) -> str:
    s = raw.strip()
    low = s.lower()
    if low == "q":
        raise NavCancel()
    if low == "h":
        raise NavHome()
    if low == "b":
        raise NavBack()
    return s


def browse(
    start: Path | None,
    *,
    mode: str,
    exts: set[str] | None,
    title: str,
) -> Path | list[Path]:
    """
    mode: file | files | folder
    """
    current = start if start and start.exists() else home_path()
    selected: list[Path] = []
    page = 0
    while True:
        clear()
        banner(f"📍 {title}")
        print(f"{C.CYAN}Current:{C.RESET} {current}")
        print()
        box_top()
        box_center("FILE BROWSER", C.BOLD + C.BCYAN)
        box_mid()
        try:
            if mode == "folder":
                folders, _ignored = list_dir(current, ALL_EXTS)
                files = []
            else:
                folders, files = list_dir(current, exts or ALL_EXTS)
        except PermissionError as exc:
            box_row(" ❌ " + str(exc), C.BRED)
            folders, files = [], []
        except OSError as exc:
            box_row(" ❌ " + str(exc), C.BRED)
            folders, files = [], []

        items: list[tuple[str, Path | None, str]] = []
        parent = current.parent if current.parent != current else None
        if parent is not None:
            items.append(("folder", parent, ".."))
        for f in folders:
            items.append(("folder", f, f.name))
        for f in files:
            items.append(("file", f, f.name))

        if not items:
            box_row(" (empty / no supported files)", C.DIM)

        total_pages = max(1, (len(items) + PAGE_SIZE - 1) // PAGE_SIZE)
        page = max(0, min(page, total_pages - 1))
        start_i = page * PAGE_SIZE
        view = items[start_i : start_i + PAGE_SIZE]
        for idx, (kind, path, label) in enumerate(view, start=1):
            mark = ""
            if kind == "file" and path in selected:
                mark = "✓ "
            if kind == "folder":
                box_row(f" {idx}. 📁 {label}", C.BBLUE)
            else:
                try:
                    sz = format_size(path.stat().st_size) if path else ""
                except OSError:
                    sz = ""
                color = C.BGREEN if mark else C.WHITE
                box_row(f" {idx}. {mark}📄 {label}  {sz}", color)
        box_mid()
        box_row(f" Page {page + 1}/{total_pages}   items: {len(items)}", C.DIM)
        if mode == "files" and selected:
            box_row(f" Selected: {len(selected)} file(s)", C.BGREEN)
        box_bot()
        print(f"{C.BGREEN}Tip:{C.RESET} file ka number dabao — turant select ho jayegi")
        print(f"{C.YELLOW}B{C.RESET}=Back  {C.YELLOW}H{C.RESET}=Home  {C.YELLOW}Q{C.RESET}=Cancel  {C.YELLOW}N/P{C.RESET}=Next/Prev page")
        if mode == "folder":
            print(f"{C.YELLOW}Enter{C.RESET} ya {C.YELLOW}S{C.RESET}= isi folder ko choose karo   {C.YELLOW}C{C.RESET}=naya folder")
        if mode == "files":
            print(f"{C.YELLOW}1,2,3{C.RESET}=ek se zyada files   phir {C.YELLOW}D{C.RESET}=Done")

        raw = ask(f"\n{C.YELLOW}Choice: {C.RESET}")
        try:
            raw = handle_nav(raw)
        except NavHome:
            current = home_path()
            page = 0
            continue
        except NavBack:
            if current.parent != current:
                current = current.parent
                page = 0
            continue
        except NavCancel:
            raise

        low = raw.lower()
        if mode == "folder" and raw == "":
            return current
        if low == "n":
            page = min(total_pages - 1, page + 1)
            continue
        if low == "p":
            page = max(0, page - 1)
            continue
        if mode == "folder" and low == "s":
            return current
        if mode == "folder" and low == "c":
            name = safe_filename(ask(f"{C.YELLOW}New folder name: {C.RESET}"))
            if not name:
                continue
            target = current / name
            try:
                target.mkdir(parents=True, exist_ok=False)
                current = target
                page = 0
            except FileExistsError:
                error_box("Folder already exists.")
                pause()
            except OSError as exc:
                error_box(str(exc))
                pause()
            continue
        if mode == "files" and low == "d":
            if selected:
                uniq = []
                seen = set()
                for p in selected:
                    if p not in seen:
                        seen.add(p)
                        uniq.append(p)
                return uniq
            error_box("Koi file select nahi hui.")
            pause()
            continue

        nums = []
        ok_multi = True
        if mode == "files" and "," in raw:
            for part in raw.split(","):
                part = part.strip()
                if not part.isdigit():
                    ok_multi = False
                    break
                nums.append(int(part))
        elif raw.isdigit():
            nums = [int(raw)]
        else:
            ok_multi = False

        if not ok_multi or not nums:
            continue

        picked_files: list[Path] = []
        opened_folder = False
        for n in nums:
            if n < 1 or n > len(view):
                continue
            kind, path, _label = view[n - 1]
            if kind == "folder" and path is not None:
                current = path
                page = 0
                opened_folder = True
                break
            if kind == "file" and path is not None:
                if mode == "file":
                    return path
                picked_files.append(path)

        if opened_folder:
            continue

        if picked_files:
            if len(nums) == 1:
                return [picked_files[0]]
            for path in picked_files:
                if path in selected:
                    selected.remove(path)
                else:
                    selected.append(path)


def default_start_folder() -> Path:
    for p in (
        Path("/sdcard/Download"),
        Path("/storage/emulated/0/Download"),
        Path.home() / "storage" / "downloads",
        Path.home() / "storage" / "shared" / "Download",
    ):
        if p.exists() and os.access(p, os.R_OK):
            return p
    return home_path()


def pick_source_files(fmt: str) -> list[Path]:
    start = default_start_folder()
    result = browse(
        start,
        mode="files",
        exts=FORMAT_EXTS[fmt],
        title=f"SELECT SOURCE  ({FORMAT_LABEL[fmt]})",
    )
    if isinstance(result, Path):
        return [result]
    return result


def pick_output_folder() -> Path:
    start = SETTINGS.get("default_output") or None
    start_path = Path(start) if start else home_path()
    if not start_path.exists():
        start_path = home_path()
    result = browse(
        start_path,
        mode="folder",
        exts=None,
        title="SELECT OUTPUT FOLDER",
    )
    if isinstance(result, list):
        raise RuntimeError("Output folder select nahi hua.")
    return result


def load_json_data(path: Path):
    try:
        with path.open("r", encoding="utf-8-sig") as f:
            return json.load(f)
    except json.JSONDecodeError as exc:
        raise ValueError(f"❌ INVALID JSON\n{path.name}\n{exc}") from exc
    except OSError as exc:
        raise ValueError(f"{path.name} read nahi ho paayi.\n{exc}") from exc


def tables_from_json_data(data, default_name: str) -> list[Table]:
    tables: list[Table] = []
    if isinstance(data, list):
        rows = [clean_row(x) for x in data if isinstance(x, dict)]
        tables.append(Table(safe_ident(default_name), rows=rows))
    elif isinstance(data, dict):
        list_keys = [k for k, v in data.items() if isinstance(v, list)]
        if list_keys:
            for k in list_keys:
                rows = [clean_row(x) for x in data[k] if isinstance(x, dict)]
                tables.append(Table(safe_ident(str(k)), rows=rows))
        else:
            row = clean_row(data)
            tables.append(Table(safe_ident(default_name), rows=[row] if row else []))
    else:
        raise ValueError("JSON array of objects ya object/table structure honi chahiye.")
    for t in tables:
        t.finalize()
    return [t for t in tables if t.columns]


def load_json(paths: list[Path]) -> list[Table]:
    merged: dict[str, Table] = {}
    for path in paths:
        data = load_json_data(path)
        for table in tables_from_json_data(data, path.stem):
            key = table.name
            if key not in merged:
                merged[key] = Table(key, columns=list(table.columns), rows=[])
            merged[key].rows.extend(table.rows)
            for col in table.columns:
                if col not in merged[key].columns:
                    merged[key].columns.append(col)
    out = list(merged.values())
    for t in out:
        t.finalize()
    if not out:
        raise ValueError("JSON se koi usable row nahi mili.")
    return out


def load_csv(paths: list[Path]) -> list[Table]:
    merged: dict[str, Table] = {}
    for path in paths:
        try:
            f = path.open("r", encoding="utf-8-sig", newline="")
        except OSError as exc:
            raise ValueError(f"{path.name} read nahi ho paayi.\n{exc}") from exc
        with f:
            reader = csv.DictReader(f)
            if not reader.fieldnames or not any((h or "").strip() for h in reader.fieldnames):
                raise ValueError(f"{path.name} empty hai ya header nahi hai.")
            name = safe_ident(path.stem)
            table = merged.setdefault(name, Table(name, columns=[], rows=[]))
            for raw in reader:
                row = {}
                for k, v in raw.items():
                    if k is None:
                        continue
                    row[str(k).strip()] = "" if v is None else v
                cleaned = clean_row(row)
                if cleaned:
                    table.rows.append(cleaned)
    out = list(merged.values())
    for t in out:
        t.finalize()
    if not out or all(t.count == 0 for t in out):
        raise ValueError("CSV se koi usable row nahi mili.")
    return out


def load_xlsx(paths: list[Path]) -> list[Table]:
    openpyxl = ensure_openpyxl()
    merged: dict[str, Table] = {}
    for path in paths:
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        except Exception as exc:
            raise ValueError(f"{path.name} invalid XLSX hai.\n{exc}") from exc
        try:
            sheet_names = list(wb.sheetnames)
            if "Data" in sheet_names:
                chosen = ["Data"]
            else:
                chosen = sheet_names
            if not chosen:
                raise ValueError(f"{path.name} mein worksheet nahi hai.")
            for sheet_name in chosen:
                ws = wb[sheet_name]
                rows_iter = ws.iter_rows(values_only=True)
                try:
                    header_row = next(rows_iter)
                except StopIteration:
                    continue
                headers = []
                for i, h in enumerate(header_row or []):
                    title = str(h).strip() if h is not None else ""
                    headers.append(title or f"col_{i + 1}")
                if not any(headers):
                    continue
                tname = safe_ident(sheet_name if sheet_name != "Data" else path.stem)
                table = merged.setdefault(tname, Table(tname, columns=[], rows=[]))
                for raw in rows_iter:
                    row = {}
                    empty = True
                    for i, h in enumerate(headers):
                        val = raw[i] if raw and i < len(raw) else None
                        nv = normalize_cell(val)
                        if not is_empty_value(nv):
                            row[h] = nv
                            empty = False
                    if not empty:
                        table.rows.append(row)
        finally:
            wb.close()
    out = list(merged.values())
    for t in out:
        t.finalize()
    if not out or all(t.count == 0 for t in out):
        raise ValueError("XLSX se koi usable row nahi mili.")
    return out


def sqlite_tables(conn: sqlite3.Connection) -> list[str]:
    cur = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
    )
    return [r[0] for r in cur.fetchall()]


def load_sqlite_tables_from_conn(conn: sqlite3.Connection, names: list[str]) -> list[Table]:
    out = []
    for name in names:
        try:
            cur = conn.execute(f"SELECT * FROM {quote_ident(name)}")
        except sqlite3.Error as exc:
            raise ValueError(f"Table '{name}' read fail.\n{exc}") from exc
        cols = [d[0] for d in cur.description or []]
        table = Table(safe_ident(name), columns=cols, rows=[])
        while True:
            chunk = cur.fetchmany(1000)
            if not chunk:
                break
            for rec in chunk:
                row = {}
                for i, col in enumerate(cols):
                    nv = normalize_cell(rec[i])
                    if not is_empty_value(nv):
                        row[col] = nv
                if row:
                    table.rows.append(row)
        table.finalize()
        out.append(table)
    if not out:
        raise ValueError("Selected table empty / missing hai.")
    return out


def choose_sqlite_tables(names: list[str]) -> list[str]:
    if not names:
        raise ValueError("Database mein koi user table nahi hai.")
    if len(names) == 1:
        return names
    while True:
        clear()
        banner("💾 SQLITE TABLE SELECT")
        box_top()
        box_center("SQLITE TABLE SELECT", C.BOLD + C.BMAGENTA)
        box_mid()
        for i, name in enumerate(names, start=1):
            box_row(f" {i}. {name}", C.WHITE)
        box_mid()
        box_row(" A = All tables", C.YELLOW)
        box_row(" Example: 1   or   1,3", C.DIM)
        box_bot()
        raw = ask(f"\n{C.YELLOW}Select table(s): {C.RESET}")
        try:
            raw = handle_nav(raw)
        except (NavCancel, NavHome, NavBack):
            raise NavCancel()
        if raw.lower() == "a":
            return names
        parts = [p.strip() for p in raw.split(",") if p.strip()]
        picked = []
        ok = True
        for p in parts:
            if not p.isdigit() or not (1 <= int(p) <= len(names)):
                ok = False
                break
            picked.append(names[int(p) - 1])
        if ok and picked:
            uniq = []
            for n in picked:
                if n not in uniq:
                    uniq.append(n)
            return uniq
        error_box("Invalid table selection.")
        pause()


def load_sqlite(paths: list[Path]) -> list[Table]:
    merged: dict[str, Table] = {}
    for path in paths:
        try:
            conn = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
        except sqlite3.Error:
            try:
                conn = sqlite3.connect(str(path))
            except sqlite3.Error as exc:
                raise ValueError(f"{path.name} invalid SQLite database hai.\n{exc}") from exc
        try:
            names = sqlite_tables(conn)
            chosen = choose_sqlite_tables(names)
            tables = load_sqlite_tables_from_conn(conn, chosen)
            for t in tables:
                if t.name not in merged:
                    merged[t.name] = Table(t.name, columns=list(t.columns), rows=[])
                merged[t.name].rows.extend(t.rows)
                for c in t.columns:
                    if c not in merged[t.name].columns:
                        merged[t.name].columns.append(c)
        finally:
            conn.close()
    out = list(merged.values())
    for t in out:
        t.finalize()
    if not out:
        raise ValueError("SQLite se data nahi mila.")
    return out


def split_sql_statements(script: str) -> list[str]:
    stmts: list[str] = []
    buf: list[str] = []
    i = 0
    n = len(script)
    in_s = in_d = False
    while i < n:
        ch = script[i]
        nxt = script[i + 1] if i + 1 < n else ""
        if not in_s and not in_d and ch == "-" and nxt == "-":
            while i < n and script[i] != "\n":
                i += 1
            continue
        if not in_s and not in_d and ch == "/" and nxt == "*":
            i += 2
            while i < n - 1 and not (script[i] == "*" and script[i + 1] == "/"):
                i += 1
            i += 2
            continue
        buf.append(ch)
        if ch == "'" and not in_d:
            if in_s and nxt == "'":
                buf.append(nxt)
                i += 2
                continue
            in_s = not in_s
        elif ch == '"' and not in_s:
            in_d = not in_d
        elif ch == ";" and not in_s and not in_d:
            stmt = "".join(buf).strip()
            if stmt and stmt != ";":
                stmts.append(stmt)
            buf = []
        i += 1
    tail = "".join(buf).strip()
    if tail:
        stmts.append(tail if tail.endswith(";") else tail + ";")
    return stmts


def load_sql(paths: list[Path]) -> list[Table]:
    tmp = tempfile.NamedTemporaryFile(prefix="rishi_", suffix=".db", delete=False)
    tmp.close()
    db_path = tmp.name
    try:
        conn = sqlite3.connect(db_path)
        conn.execute("PRAGMA foreign_keys=OFF")
        try:
            for path in paths:
                try:
                    text = path.read_text(encoding="utf-8-sig")
                except OSError as exc:
                    raise ValueError(f"{path.name} read nahi ho paayi.\n{exc}") from exc
                if not text.strip():
                    raise ValueError(f"{path.name} empty SQL file hai.")
                try:
                    conn.executescript(text)
                except sqlite3.Error:
                    stmts = split_sql_statements(text)
                    if not stmts:
                        raise ValueError(f"{path.name} mein SQL statement nahi mili.")
                    for stmt in stmts:
                        try:
                            conn.execute(stmt)
                        except sqlite3.Error as exc:
                            raise ValueError(f"{path.name} SQL error:\n{exc}\n---\n{stmt[:180]}") from exc
                conn.commit()
            names = sqlite_tables(conn)
            if not names:
                raise ValueError("SQL import ke baad koi table nahi bani. CREATE/INSERT check karo.")
            chosen = choose_sqlite_tables(names)
            tables = load_sqlite_tables_from_conn(conn, chosen)
        finally:
            conn.close()
    finally:
        try:
            os.unlink(db_path)
        except OSError:
            pass
    return tables


LOADERS = {
    "json": load_json,
    "csv": load_csv,
    "xlsx": load_xlsx,
    "sqlite": load_sqlite,
    "sql": load_sql,
}


class SplitJsonWriter:
    def __init__(self, folder: Path, stem: str, max_bytes: int | None):
        self.folder = folder
        self.stem = stem
        self.max_bytes = max_bytes
        self.part = 0
        self.count_in_file = 0
        self.total = 0
        self.bytes_written = 0
        self.f = None
        self.files: list[Path] = []
        self.path: Path | None = None
        self._open_new()

    def _open_new(self) -> None:
        if self.f is not None:
            self.f.write("\n]\n")
            self.f.close()
        self.path = alloc_path(self.folder, self.stem, ".json", None if not self.files else None)
        # alloc_path already skips existing; first file uses next free name in series
        if not self.files:
            self.path = alloc_path(self.folder, self.stem, ".json")
        else:
            # continue numbered sequence after last allocated stem index
            self.path = alloc_path(self.folder, self.stem, ".json")
        self.f = self.path.open("w", encoding="utf-8")
        self.f.write("[\n")
        self.bytes_written = 2
        self.count_in_file = 0
        self.files.append(self.path)

    def add(self, row: dict) -> None:
        dumped = json.dumps(row, ensure_ascii=False, indent=2)
        indented = "  " + dumped.replace("\n", "\n  ")
        chunk = indented if self.count_in_file == 0 else ",\n" + indented
        extra = len(chunk.encode("utf-8"))
        if (
            self.max_bytes
            and self.count_in_file > 0
            and self.bytes_written + extra + 2 > self.max_bytes
        ):
            self.part += 1
            self._open_new()
            chunk = indented
            extra = len(chunk.encode("utf-8"))
        self.f.write(chunk)
        self.bytes_written += extra
        self.count_in_file += 1
        self.total += 1
        if self.total % PROGRESS_EVERY == 0:
            print(
                f"\r  🔄 rows={self.total:,}  file={self.path.name}  size={format_size(self.bytes_written)}",
                end="",
                flush=True,
            )

    def close(self) -> None:
        if self.f is None:
            return
        if self.count_in_file == 0:
            self.f.close()
            self.path.write_text("[]\n", encoding="utf-8")
        else:
            self.f.write("\n]\n")
            self.f.close()
        self.f = None
        if self.total >= PROGRESS_EVERY:
            print()


def save_json(tables: list[Table], folder: Path, stem: str, max_bytes: int | None) -> list[Path]:
    if len(tables) == 1:
        writer = SplitJsonWriter(folder, stem, max_bytes)
        for row in tables[0].rows:
            writer.add(clean_row(row))
        writer.close()
        return writer.files
    # multi-table: one object file, split only if single requested? split by table files
    if max_bytes is None:
        path = alloc_path(folder, stem, ".json")
        payload = {t.name: [clean_row(r) for r in t.rows] for t in tables}
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return [path]
    files = []
    for t in tables:
        writer = SplitJsonWriter(folder, f"{stem}_{t.name}", max_bytes)
        for row in t.rows:
            writer.add(clean_row(row))
        writer.close()
        files.extend(writer.files)
    return files


def save_csv(tables: list[Table], folder: Path, stem: str, max_bytes: int | None) -> list[Path]:
    files: list[Path] = []
    for t_index, table in enumerate(tables):
        tstem = stem if len(tables) == 1 else f"{stem}_{table.name}"
        headers = table.columns or unique_columns(table.rows)
        f = None
        current = None
        written = 0
        count = 0

        def open_new():
            nonlocal f, current, written, count
            if f:
                f.close()
            current = alloc_path(folder, tstem, ".csv")
            f = current.open("w", encoding="utf-8-sig", newline="")
            writer = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
            writer.writeheader()
            written = f.tell()
            count = 0
            files.append(current)
            return writer

        writer = open_new()
        for i, row in enumerate(table.rows, start=1):
            clean = {h: row.get(h, "") for h in headers}
            # estimate
            line_est = sum(len(str(v)) + 1 for v in clean.values()) + 2
            if max_bytes and count > 0 and written + line_est > max_bytes:
                writer = open_new()
            writer.writerow(clean)
            count += 1
            try:
                written = f.tell()
            except OSError:
                written += line_est
            if i % PROGRESS_EVERY == 0:
                print(
                    f"\r  🔄 rows={i:,}  file={current.name}  size={format_size(written)}",
                    end="",
                    flush=True,
                )
        if f:
            f.close()
        if table.count >= PROGRESS_EVERY:
            print()
    return files


def save_xlsx(tables: list[Table], folder: Path, stem: str, max_bytes: int | None) -> list[Path]:
    openpyxl = ensure_openpyxl()
    files: list[Path] = []

    def new_wb():
        wb = openpyxl.Workbook(write_only=True)
        return wb

    if not max_bytes:
        path = alloc_path(folder, stem, ".xlsx")
        wb = new_wb()
        for table in tables:
            ws = wb.create_sheet(table.name[:31] if table.name else "Data")
            headers = table.columns or unique_columns(table.rows)
            ws.append(headers)
            for i, row in enumerate(table.rows, start=1):
                ws.append([row.get(h, "") for h in headers])
                if i % PROGRESS_EVERY == 0:
                    print(f"\r  🔄 rows={i:,}  file={path.name}", end="", flush=True)
        # remove default write_only empty? write_only has no default sheet
        if tables and "Sheet" in getattr(wb, "sheetnames", []):
            pass
        wb.save(path)
        files.append(path)
        if any(t.count >= PROGRESS_EVERY for t in tables):
            print()
        return files

    # Split rows across multiple workbooks. Size is estimated from cell text.
    for table in tables:
        tstem = stem if len(tables) == 1 else f"{stem}_{table.name}"
        headers = table.columns or unique_columns(table.rows)
        wb = new_wb()
        ws = wb.create_sheet("Data")
        ws.append(headers)
        est = 64
        count = 0
        total = 0
        for row in table.rows:
            values = [row.get(h, "") for h in headers]
            row_est = sum(len(str(v)) + 2 for v in values) + 8
            if count > 0 and est + row_est > max_bytes:
                path = alloc_path(folder, tstem, ".xlsx")
                wb.save(path)
                files.append(path)
                wb = new_wb()
                ws = wb.create_sheet("Data")
                ws.append(headers)
                est = 64
                count = 0
            ws.append(values)
            est += row_est
            count += 1
            total += 1
            if total % PROGRESS_EVERY == 0:
                print(f"\r  🔄 rows={total:,}  table={table.name}", end="", flush=True)
        path = alloc_path(folder, tstem, ".xlsx")
        wb.save(path)
        files.append(path)
        if total >= PROGRESS_EVERY:
            print()
    return files


def write_sqlite_db(tables: list[Table], path: Path) -> None:
    if path.exists():
        raise FileExistsError(str(path))
    conn = sqlite3.connect(str(path))
    try:
        conn.execute("PRAGMA journal_mode=OFF")
        for table in tables:
            name = safe_ident(table.name)
            headers = table.columns or unique_columns(table.rows)
            if not headers:
                continue
            cols_sql = []
            for col in headers:
                cols_sql.append(f"{quote_ident(col)} {infer_sql_type(col, table.rows)}")
            conn.execute(f"CREATE TABLE {quote_ident(name)} ({', '.join(cols_sql)})")
            placeholders = ",".join(["?"] * len(headers))
            insert_sql = f"INSERT INTO {quote_ident(name)} ({', '.join(quote_ident(c) for c in headers)}) VALUES ({placeholders})"
            batch = []
            total = 0
            for row in table.rows:
                batch.append([row.get(h) for h in headers])
                if len(batch) >= BATCH_INSERT:
                    conn.executemany(insert_sql, batch)
                    total += len(batch)
                    batch = []
                    if total % PROGRESS_EVERY == 0:
                        print(f"\r  🔄 rows={total:,}  table={name}", end="", flush=True)
            if batch:
                conn.executemany(insert_sql, batch)
                total += len(batch)
            conn.commit()
        if any(t.count >= PROGRESS_EVERY for t in tables):
            print()
    finally:
        conn.close()


def save_sqlite(tables: list[Table], folder: Path, stem: str, max_bytes: int | None) -> list[Path]:
    if not max_bytes:
        path = alloc_path(folder, stem, ".db")
        write_sqlite_db(tables, path)
        return [path]
    print(f"{C.YELLOW}⚠️  SQLite .db ko raw size-cut nahi karte.{C.RESET}")
    print(f"{C.YELLOW}   Split = same schema ke multiple valid .db files (row chunks).{C.RESET}")
    files = []
    # split first table rows across dbs; extra tables go fully into each? Better:
    # chunk ALL tables' primary (or each table independently).
    # Practical: chunk the largest / first table, keep others only in first file if small.
    # Spec: row/data export splitting. Chunk every table independently into numbered dbs
    # with one table each when multiple, or chunk single table.
    if len(tables) == 1:
        table = tables[0]
        headers = table.columns
        chunk_rows: list[dict] = []
        est = 256
        part_tables: list[list[dict]] = []
        for row in table.rows:
            row_est = sum(len(str(v)) + 4 for v in row.values())
            if chunk_rows and est + row_est > max_bytes:
                part_tables.append(chunk_rows)
                chunk_rows = []
                est = 256
            chunk_rows.append(row)
            est += row_est
        if chunk_rows:
            part_tables.append(chunk_rows)
        for chunk in part_tables:
            path = alloc_path(folder, stem, ".db")
            write_sqlite_db([Table(table.name, columns=headers, rows=chunk)], path)
            files.append(path)
        return files
    # multiple tables: one db per table (valid), further chunk if needed
    for table in tables:
        files.extend(save_sqlite([table], folder, f"{stem}_{table.name}", max_bytes))
    return files


def table_to_sql_script(table: Table) -> str:
    name = safe_ident(table.name)
    headers = table.columns or unique_columns(table.rows)
    lines = []
    cols_sql = ", ".join(
        f"{quote_ident(col)} {infer_sql_type(col, table.rows)}" for col in headers
    )
    lines.append(f"CREATE TABLE IF NOT EXISTS {quote_ident(name)} ({cols_sql});")
    col_list = ", ".join(quote_ident(c) for c in headers)
    for row in table.rows:
        vals = ", ".join(sql_literal(row.get(h)) for h in headers)
        lines.append(f"INSERT INTO {quote_ident(name)} ({col_list}) VALUES ({vals});")
    return "\n".join(lines) + "\n"


def save_sql(tables: list[Table], folder: Path, stem: str, max_bytes: int | None) -> list[Path]:
    statements: list[str] = []
    for table in tables:
        name = safe_ident(table.name)
        headers = table.columns or unique_columns(table.rows)
        cols_sql = ", ".join(
            f"{quote_ident(col)} {infer_sql_type(col, table.rows)}" for col in headers
        )
        statements.append(f"CREATE TABLE IF NOT EXISTS {quote_ident(name)} ({cols_sql});")
        col_list = ", ".join(quote_ident(c) for c in headers)
        for row in table.rows:
            vals = ", ".join(sql_literal(row.get(h)) for h in headers)
            statements.append(f"INSERT INTO {quote_ident(name)} ({col_list}) VALUES ({vals});")

    files: list[Path] = []
    if not max_bytes:
        path = alloc_path(folder, stem, ".sql")
        path.write_text("\n".join(statements) + "\n", encoding="utf-8")
        return [path]

    buf: list[str] = []
    size = 0
    for stmt in statements:
        piece = stmt + "\n"
        extra = len(piece.encode("utf-8"))
        if buf and max_bytes and size + extra > max_bytes:
            path = alloc_path(folder, stem, ".sql")
            path.write_text("".join(buf), encoding="utf-8")
            files.append(path)
            buf = []
            size = 0
        buf.append(piece)
        size += extra
    if buf:
        path = alloc_path(folder, stem, ".sql")
        path.write_text("".join(buf), encoding="utf-8")
        files.append(path)
    return files


SAVERS = {
    "json": save_json,
    "csv": save_csv,
    "xlsx": save_xlsx,
    "sqlite": save_sqlite,
    "sql": save_sql,
}


def ask_output_name(default: str) -> str:
    print(f"{C.CYAN}Default name:{C.RESET} {default}")
    raw = ask(f"{C.YELLOW}Output name (Enter = default): {C.RESET}")
    try:
        raw = handle_nav(raw)
    except (NavCancel, NavHome, NavBack):
        raise NavCancel()
    return safe_filename(raw or default)


def ask_table_names(tables: list[Table]) -> None:
    print()
    box_top()
    box_center("SQLITE TABLE NAME", C.BOLD + C.MAGENTA)
    box_mid()
    box_row(" Import ke liye table name set karo.", C.WHITE)
    box_bot()
    used = set()
    for table in tables:
        raw = ask(f"{C.YELLOW}Table name [{table.name}]: {C.RESET}")
        try:
            raw = handle_nav(raw)
        except (NavCancel, NavHome, NavBack):
            raise NavCancel()
        name = safe_ident(raw or table.name)
        base = name
        n = 2
        while name.lower() in used:
            name = f"{base}_{n}"
            n += 1
        used.add(name.lower())
        table.name = name


def ask_split() -> int | None:
    print()
    box_top()
    box_center("SPLIT MODE", C.BOLD + C.MAGENTA)
    box_mid()
    box_row("  1. SINGLE", C.WHITE)
    box_row("  2. SPLIT", C.WHITE)
    box_bot()
    raw = ask(f"\n{C.YELLOW}Select: {C.RESET}")
    try:
        raw = handle_nav(raw)
    except (NavCancel, NavHome, NavBack):
        raise NavCancel()
    if raw != "2":
        return None
    print(f"{C.DIM}Examples: 100 KB   1 MB   10 MB   100 MB{C.RESET}")
    default = SETTINGS.get("default_split") or "1 MB"
    raw = ask(f"{C.YELLOW}Maximum size [{default}]: {C.RESET}") or default
    try:
        raw = handle_nav(raw)
    except (NavCancel, NavHome, NavBack):
        raise NavCancel()
    return parse_size(raw)


def convert_flow(src_fmt: str) -> None:
    targets = [f for f in FORMATS if f != src_fmt]
    clear()
    banner(f"SOURCE FORMAT : {FORMAT_LABEL[src_fmt]}")
    box_top()
    box_center("OUTPUT FORMAT", C.BOLD + C.MAGENTA)
    box_mid()
    for i, fmt in enumerate(targets, start=1):
        box_row(f"  {i}. {FORMAT_LABEL[fmt]}", C.WHITE)
    box_mid()
    box_row(" B=Back  H=Home  Q=Cancel", C.DIM)
    box_bot()
    raw = ask(f"\n{C.YELLOW}Select: {C.RESET}")
    try:
        raw = handle_nav(raw)
    except NavCancel:
        return
    except NavHome:
        return
    except NavBack:
        return
    if not raw.isdigit() or not (1 <= int(raw) <= len(targets)):
        error_box("Invalid output format.")
        pause()
        return
    dst_fmt = targets[int(raw) - 1]

    try:
        sources = pick_source_files(src_fmt)
    except NavCancel:
        return
    except NavHome:
        return
    except NavBack:
        return
    except Exception as exc:
        error_box(str(exc))
        pause()
        return

    if not sources:
        error_box("Koi source file select nahi hui.")
        pause()
        return

    print()
    print(f"{C.BGREEN}✓ Selected:{C.RESET} {', '.join(p.name for p in sources[:5])}")
    if len(sources) > 5:
        print(f"{C.DIM}  +{len(sources) - 5} more{C.RESET}")

    try:
        tables = LOADERS[src_fmt](sources)
    except NavCancel:
        return
    except Exception as exc:
        error_box(str(exc))
        pause()
        return

    source_dir = sources[0].parent
    print()
    box_top()
    box_center("OUTPUT FOLDER", C.BOLD + C.MAGENTA)
    box_mid()
    box_row("  1. Same folder (easy)", C.BGREEN)
    box_row(f"     {source_dir}", C.DIM)
    box_row("  2. Dusra folder choose karo", C.WHITE)
    box_bot()
    raw = ask(f"\n{C.YELLOW}Select [1]: {C.RESET}") or "1"
    try:
        raw = handle_nav(raw)
    except (NavCancel, NavHome, NavBack):
        return
    if raw == "2":
        try:
            out_dir = pick_output_folder()
        except (NavCancel, NavHome, NavBack):
            return
        except Exception as exc:
            error_box(str(exc))
            pause()
            return
    else:
        out_dir = source_dir

    default_name = sources[0].stem if sources else "converted"
    try:
        print()
        print(f"{C.CYAN}Folder + file name default:{C.RESET} {default_name}/")
        raw_name = ask(f"{C.YELLOW}Naya naam (Enter = same): {C.RESET}")
        raw_name = handle_nav(raw_name)
        stem = safe_filename(raw_name or default_name)
        if dst_fmt == "sqlite" and len(tables) == 1:
            tables[0].name = safe_ident(stem)
        elif dst_fmt == "sqlite":
            ask_table_names(tables)
        print()
        print(f"{C.YELLOW}1{C.RESET}=SINGLE file   {C.YELLOW}2{C.RESET}=SPLIT")
        raw_split = ask(f"{C.YELLOW}Mode [1]: {C.RESET}") or "1"
        raw_split = handle_nav(raw_split)
        max_bytes = None
        if raw_split == "2":
            default = SETTINGS.get("default_split") or "1 MB"
            size_raw = ask(f"{C.YELLOW}Max size [{default}]: {C.RESET}") or default
            size_raw = handle_nav(size_raw)
            max_bytes = parse_size(size_raw)
    except NavCancel:
        return
    except ValueError as exc:
        error_box(str(exc))
        pause()
        return

    parent_dir = out_dir
    try:
        out_dir = get_unique_folder(parent_dir, stem)
    except OSError as exc:
        error_box(f"Folder nahi ban paya.\n{exc}")
        pause()
        return

    clear()
    banner("🔄 CONVERTING")
    box_top()
    box_center("🔄 CONVERTING", C.BOLD + C.BYELLOW)
    box_mid()
    box_row(f" Source : {FORMAT_LABEL[src_fmt]} → {FORMAT_LABEL[dst_fmt]}", C.WHITE)
    box_row(f" Files  : {len(sources)}", C.WHITE)
    box_row(f" Output : {out_dir}", C.WHITE)
    box_row(f" Mode   : {'SPLIT ' + format_size(max_bytes) if max_bytes else 'SINGLE'}", C.WHITE)
    box_bot()
    print()

    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        files = SAVERS[dst_fmt](tables, out_dir, stem, max_bytes)
    except PermissionError:
        error_box("Permission denied.\nAgar Android storage chahiye to:\ntermux-setup-storage")
        pause()
        return
    except FileExistsError as exc:
        error_box(f"File already exists:\n{exc}")
        pause()
        return
    except Exception as exc:
        error_box(str(exc))
        pause()
        return

    rows = sum(t.count for t in tables)
    cols = max((len(t.columns) for t in tables), default=0)
    success_box(rows, cols, files, out_dir)
    pause()


def settings_menu() -> None:
    global SETTINGS
    while True:
        clear()
        banner("SETTINGS")
        box_top()
        box_center("SETTINGS", C.BOLD + C.MAGENTA)
        box_mid()
        colors = "ON" if SETTINGS.get("colors", True) else "OFF"
        dout = SETTINGS.get("default_output") or "(not set)"
        dsplit = SETTINGS.get("default_split") or "1 MB"
        box_row(f"  1. Color              : {colors}", C.WHITE)
        box_row("  2. Show unsupported   : NEVER", C.DIM)
        box_row(f"  3. Default output     : {dout}", C.WHITE)
        box_row(f"  4. Default split size : {dsplit}", C.WHITE)
        box_row("  5. Reset settings", C.WHITE)
        box_row("  6. Back", C.WHITE)
        box_bot()
        raw = ask(f"\n{C.YELLOW}Select: {C.RESET}")
        try:
            raw = handle_nav(raw)
        except (NavCancel, NavHome, NavBack):
            return
        if raw == "1":
            SETTINGS["colors"] = not bool(SETTINGS.get("colors", True))
            apply_color_mode(SETTINGS["colors"])
            save_settings(SETTINGS)
        elif raw == "2":
            pause("Unsupported files picker mein kabhi nahi dikhte.")
        elif raw == "3":
            try:
                folder = pick_output_folder()
                SETTINGS["default_output"] = str(folder)
                save_settings(SETTINGS)
            except (NavCancel, NavHome, NavBack):
                pass
            except Exception as exc:
                error_box(str(exc))
                pause()
        elif raw == "4":
            raws = ask(f"{C.YELLOW}Default split size: {C.RESET}")
            try:
                parse_size(raws)
                SETTINGS["default_split"] = raws.strip()
                save_settings(SETTINGS)
            except ValueError as exc:
                error_box(str(exc))
                pause()
        elif raw == "5":
            SETTINGS = default_settings()
            apply_color_mode(True)
            save_settings(SETTINGS)
        elif raw == "6":
            return


def main_menu() -> None:
    check_storage()
    while True:
        clear()
        banner()
        box_top()
        box_center("MAIN MENU", C.BOLD + C.MAGENTA)
        box_mid()
        box_row("  1. JSON", C.WHITE)
        box_row("  2. CSV", C.WHITE)
        box_row("  3. EXCEL (XLSX)", C.WHITE)
        box_row("  4. SQLITE", C.WHITE)
        box_row("  5. SQL", C.WHITE)
        box_row("  6. SETTINGS", C.WHITE)
        box_row("  7. EXIT", C.WHITE)
        box_bot()
        print(f"{C.DIM}  Source format choose karo, phir output format.{C.RESET}")
        raw = ask(f"\n{C.YELLOW}Select: {C.RESET}")
        mapping = {"1": "json", "2": "csv", "3": "xlsx", "4": "sqlite", "5": "sql"}
        if raw == "6":
            settings_menu()
            continue
        if raw == "7" or raw.lower() == "q":
            print(f"\n{C.BGREEN}Bye Rishi. ✅{C.RESET}\n")
            return
        if raw.lower() == "h":
            continue
        if raw in mapping:
            try:
                convert_flow(mapping[raw])
            except KeyboardInterrupt:
                print()
                continue
            except NavCancel:
                continue
            except Exception as exc:
                error_box(str(exc))
                pause()
            continue


def main() -> None:
    try:
        main_menu()
    except KeyboardInterrupt:
        print(f"\n\n{C.YELLOW}Cancelled.{C.RESET}\n")
        sys.exit(0)


if __name__ == "__main__":
    main()
