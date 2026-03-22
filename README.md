# Folder File Table → Word (Windows + macOS)

Pick a folder, generate a table for every file inside it, then **paste into Word** or export a Word document.

**Repo check:** If you see this line on GitHub, your latest push worked. *(Updated 2026-02-26.)*

## Columns generated

- File name
- Last 2 Paths (last two folder names of the file’s full path)
- Inner path (path inside the selected folder; blank if file is directly in the root)
- Run (exe/dll): `V` when extension is `.exe` or `.dll`
- Data (other): `V` for everything else
- Created (local): full date + time
- Size (bytes)

## Run (simple GUI – recommended)

This version uses the built‑in Tk GUI (no Qt needed).

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python simple_gui.py
```

On Windows (PowerShell), the activate line is:

```powershell
.\.venv\Scripts\Activate.ps1
python simple_gui.py
```

## Run (advanced Qt GUI – optional, may require extra Qt setup)

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python qt_app.py
```

On Windows (PowerShell), the activate line is:

```powershell
.\.venv\Scripts\Activate.ps1
python qt_app.py
```

## Paste into Word

You have 3 options:

- **Copy (Word table)**: copies **HTML + TSV** to clipboard. In Word, paste and you should get a real table.
- **Export DOCX…**: creates a native Word file with a real table.
- **Export HTML…**: Word opens HTML and keeps the table.

## Build an app (optional)

You can package it into a standalone app/exe using PyInstaller:

```bash
python3 -m pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed qt_app.py --name FolderFileTable
```

Output will be in `dist/`.

