# 🔄 RISHI DATA CONVERTER

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=for-the-badge&logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Termux%20%7C%20Android%20%7C%20Linux%20%7C%20Windows-green?style=for-the-badge)
![Formats](https://img.shields.io/badge/Formats-JSON%20%E2%86%94%20CSV%20%E2%86%94%20XLSX%20%E2%86%94%20SQLite%20%E2%86%94%20SQL-orange?style=for-the-badge)
![License](https://img.shields.io/badge/License-MIT-purple?style=for-the-badge)

**RISHI DATA CONVERTER** is a fast, colorful, interactive terminal application built specifically for **Termux (Android)** and desktop CLIs. It converts, merges, and splits structured datasets seamlessly across **JSON, CSV, Excel (XLSX), SQLite, and SQL** formats without requiring heavy external setups.

---

## ✨ Key Features

* **5-Way Bidirectional Conversion:** Convert any supported format into any other format (`JSON ↔ CSV ↔ XLSX ↔ SQLite ↔ SQL`).
* **Smart Size-Based File Splitting:** Export data as a **Single File** or **Split** large outputs into smaller valid chunks (e.g., `100 KB`, `1 MB`, `10 MB`) across all formats—including valid multi-part SQLite `.db` files.
* **Multi-File Merging:** Select multiple files at once (`1,2,3`) in the built-in browser to merge them cleanly into a unified dataset or multi-table database.
* **Interactive Terminal File Browser:** Built-in paginated file/folder picker with Android `/sdcard` & Termux storage detection, folder creation, and instant navigation (`B` = Back, `H` = Home, `Q` = Cancel, `N/P` = Page).
* **Automatic Schema & Type Inference:** Automatically detects `INTEGER`, `REAL`, and `TEXT` column types when generating SQLite databases or SQL `CREATE TABLE` scripts.
* **Data Cleaning & Normalization:** Strips empty/null fields, trims whitespace, handles UTF-8 BOM encodings, and safely escapes SQL literals.
* **Zero-Config Dependency Auto-Installer:** Works out of the box with Python's standard library (`json`, `csv`, `sqlite3`), and offers a 1-click prompt to auto-install `openpyxl` only when working with `.xlsx` files.
* **Persistent Settings:** Customize ANSI color mode, default output directory, and default split size saved automatically to `~/.rishi_data_converter/settings.json`.

---

## 📂 Supported Formats & Extensions

| Format | Extensions | Notes |
| :--- | :--- | :--- |
| **JSON** | `.json` | Supports arrays of objects & multi-table object keys |
| **CSV** | `.csv` | UTF-8 with BOM (`utf-8-sig`) for full Excel compatibility |
| **Excel** | `.xlsx` | Streaming read/write via `openpyxl` |
| **SQLite** | `.db`, `.sqlite`, `.sqlite3` | Interactive table selector & batched `executemany` inserts |
| **SQL** | `.sql` | Executes `CREATE` / `INSERT` scripts & exports clean SQL dumps |

---

## 🚀 Installation & Usage (Termux / Linux / Windows)

### 1. Clone the Repository
```bash
git clone [https://github.com/rishi-security-labs/File-Converter.git](https://github.com/rishi-security-labs/File-Converter.git)
cd File-Converter
```

### 2. Grant Storage Permission (Termux / Android Only)
```bash
termux-setup-storage
```

### 3. Install Optional Excel Support (Recommended)
```bash
pip install openpyxl
```
*(Note: JSON, CSV, SQLite, and SQL work with zero external packages.)*

### 4. Run the Tool
```bash
python converter.py
```
*(Replace `converter.py` with your main script filename if named differently, e.g., `main.py`.)*

---

## 🧭 Navigation Shortcuts

| Key | Action |
| :--- | :--- |
| `1, 2, 3...` | Select folder or file (comma-separated for multi-select) |
| `D` | Done selecting multiple files |
| `N` / `P` | Next / Previous page in file browser |
| `B` | Go back one directory |
| `H` | Jump to Home / Storage root |
| `C` | Create a new folder inside output picker |
| `Q` | Cancel current operation / Exit |

---

## ⚙️ Configuration

User preferences are stored locally at:
```text
~/.rishi_data_converter/settings.json
```
You can toggle terminal colors, set a default export folder, or configure your preferred split chunk size directly from **Option 6 (`SETTINGS`)** in the main menu.

---

## 📄 License

This project is licensed under the **MIT License** — free to use, modify, and distribute.
