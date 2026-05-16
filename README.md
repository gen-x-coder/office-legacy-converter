# Office Legacy Converter

A GUI tool for Windows that converts entire folders of old Microsoft Office files to modern formats in one go — not one file at a time like most online converters, but whole archives including all subfolders.

Designed for anyone with a collection of `.doc`, `.ppt` or `.pps` files from the pre-2007 era who wants to make them usable in modern software without losing layout, images or tables. Point it at a folder, and it handles the rest.

Conversion is handled by the OnlyOffice `x2t` engine, which produces accurate results comparable to opening and re-saving the file in Word or PowerPoint. The tool is available in English and Dutch.

---

## What it does

- Processes entire folder trees in one run — including all subfolders, recursively
- Converts `.doc` to `.docx`, `.ppt` to `.pptx`, and `.pps` to `.pptx`
- Creates a full backup with identical folder structure and original file timestamps before converting
- Optionally removes original files after successful conversion (backup is always made first)
- Detects files case-insensitively, so `.DOC` and `.Doc` are treated the same as `.doc`
- Saves your settings between sessions (folders, language, preferences)
- Checks for and installs missing Python dependencies automatically on first run

Most online tools convert a single file per upload. This tool is built for the opposite situation: a hard drive, NAS, or archive folder with dozens or hundreds of old Office files that all need to be migrated at once.

---

## Requirements

### Python

Download and install Python 3.10 or higher from [python.org](https://www.python.org/downloads/).

During installation on Windows, check the box **"Add Python to PATH"**. Without this, the script cannot be started from the command line.

Verify the installation by opening a terminal and running:

```
python --version
```

If you see a version number, Python is ready.

### pip

pip is the Python package manager and is included with Python 3.4 and higher. Verify it is working:

```
python -m pip --version
```

If pip is missing, run:

```
python -m ensurepip --upgrade
```

The script will also attempt this automatically if pip is not found.

### OnlyOffice x2t

The converter relies on `x2t.exe`, the internal conversion engine from OnlyOffice Desktop Editors. This is not included in this repository due to its size and license. You must supply it yourself in one of two ways.

**Option A: You have OnlyOffice Desktop Editors installed**

The script automatically finds `x2t.exe` in its default installation path. No additional steps are needed.

**Option B: Standalone use without installing OnlyOffice**

Copy the `converter` folder from your OnlyOffice installation to the same directory as `doc_converter.py`:

```
C:\Program Files\ONLYOFFICE\DesktopEditors\converter\
```

The result should look like this:

```
doc_converter.py
converter/
    x2t.exe
    (various .dll files)
```

The script checks this location first, so it will work without a full OnlyOffice installation on the target machine.

---

## Installation

No installation is required beyond Python itself. Download `doc_converter.py` and place it wherever you want to run it from. On first run, the only required Python package (`customtkinter`) is installed automatically.

---

## Usage

Open a terminal in the folder containing `doc_converter.py` and run:

```
python doc_converter.py
```

On first run, the script installs missing dependencies and then starts. On subsequent runs it starts immediately.

**Steps in the application:**

1. Choose your language (English or Dutch). This choice is saved for future sessions.
2. Choose the source folder containing the old Office files.
3. Choose the backup folder where originals will be copied before conversion.
4. Click Scan. The tool lists all found files with their extension and size.
5. Optionally check the option to delete originals after successful conversion. Originals are always backed up first regardless of this setting.
6. Click the convert button and confirm the summary. The progress bar shows per-file status.

After conversion, a report shows how many files were converted successfully and which files failed, with a separate category for files where conversion succeeded but the original could not be deleted (for example because it was open in another application).

---

## Notes

- The `converter` folder from OnlyOffice is approximately 300 MB and is not included here. You must supply it from your own OnlyOffice installation.
- OnlyOffice is licensed under [AGPL-3.0](https://www.gnu.org/licenses/agpl-3.0.html). This repository does not redistribute any OnlyOffice binaries.
- Original files are never deleted unless you explicitly enable the option, and only after a successful conversion and a successful backup.
- Conversion preserves text, images, tables, fonts, styles, page breaks and layout. This is a full conversion, not a text extraction.
- The tool was tested on Windows. The conversion logic depends on `x2t.exe`, which is a Windows binary, so this tool is Windows-only.

---

## License

The Python script (`doc_converter.py`) is released under the MIT License. Do with it as you please.
