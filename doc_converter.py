"""
Office Legacy Converter
Converts old Microsoft Office files (.doc, .ppt, .pps) to modern formats via OnlyOffice x2t.

Usage:    python doc_converter.py
Requires: OnlyOffice x2t.exe in converter/ subfolder or a full OnlyOffice installation.
          All Python dependencies are installed automatically on first run.
"""

import os
import sys
import json
import shutil
import subprocess
import threading
import time
from pathlib import Path
from tkinter import filedialog, messagebox


REQUIRED_PACKAGES = ["customtkinter"]
SETTINGS_FILE = Path(__file__).parent / "settings.json"
SUPPORTED_EXTENSIONS = {'.doc', '.ppt', '.pps'}
EXTENSION_MAP = {'.doc': '.docx', '.ppt': '.pptx', '.pps': '.pptx'}


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

def load_settings() -> dict:
    try:
        if SETTINGS_FILE.exists():
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_settings(s: dict):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(s, f, indent=4)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Dependency check — must run before importing customtkinter
# ---------------------------------------------------------------------------

def ensure_pip() -> bool:
    """Verify pip is available, installing it via ensurepip if needed."""
    try:
        subprocess.run([sys.executable, "-m", "pip", "--version"],
                       check=True, capture_output=True)
        return True
    except subprocess.CalledProcessError:
        pass
    try:
        subprocess.run([sys.executable, "-m", "ensurepip", "--upgrade"],
                       check=True, capture_output=True)
        return True
    except Exception:
        return False


def install_packages(packages: list[str]) -> tuple[bool, str]:
    """Install a list of pip packages. Returns (success, error_message)."""
    try:
        kwargs = {}
        if sys.platform == "win32":
            kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW
        subprocess.run([sys.executable, "-m", "pip", "install"] + packages,
                       check=True, capture_output=True, **kwargs)
        return True, ""
    except subprocess.CalledProcessError as e:
        return False, e.stderr.decode("utf-8", errors="ignore")


def check_dependencies() -> list[str]:
    """Return a list of package names that are not yet installed."""
    missing = []
    for pkg in REQUIRED_PACKAGES:
        try:
            __import__(pkg.replace("-", "_"))
        except ImportError:
            missing.append(pkg)
    return missing


_settings = load_settings()
_missing = check_dependencies()

if _missing and not _settings.get("deps_installed"):
    if not ensure_pip():
        print("ERROR: pip is not available and could not be installed.")
        print("Please install pip manually: https://pip.pypa.io/en/stable/installation/")
        sys.exit(1)
    ok, err = install_packages(_missing)
    if not ok:
        print(f"ERROR: Could not install required packages: {err}")
        sys.exit(1)
    _settings["deps_installed"] = True
    save_settings(_settings)

import customtkinter as ctk

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

SOFT = ("#444444", "#BBBBBB")


# ---------------------------------------------------------------------------
# Translations
# ---------------------------------------------------------------------------

TRANSLATIONS = {
    "en": {
        "app_title":           "Office Legacy Converter",
        "app_subtitle":        "Converts .doc, .ppt and .pps to modern formats with full formatting via OnlyOffice.",
        "x2t_found":           "OnlyOffice x2t found: {path}",
        "x2t_not_found":       "OnlyOffice x2t NOT found. Place the converter/ folder next to this script.",
        "source_folder":       "Source folder:",
        "backup_folder":       "Backup folder:",
        "x2t_path":            "x2t path:",
        "choose":              "Choose",
        "scan":                "Scan",
        "scanning":            "Scanning...",
        "convert_btn":         "Backup and convert",
        "converting":          "Converting...",
        "delete_cb":           "Delete originals after successful conversion",
        "delete_warning":      "  The originals are safely stored in the backup folder.",
        "not_chosen":          "Not chosen",
        "x2t_not_found_short": "Not found — choose manually",
        "list_label":          "Found files  (.doc -> .docx   |   .ppt / .pps -> .pptx)",
        "col_name":            "Filename",
        "col_ext":             "Extension",
        "col_converts_to":     "Converts to",
        "col_size":            "Size",
        "no_files_found":      "No supported files found.",
        "files_found":         "{n} files found:  {summary}",
        "warn_source":         "Please choose a source folder first.",
        "warn_x2t":            "OnlyOffice x2t not found.\nUse the Choose button next to x2t path.",
        "warn_backup":         "Please choose a backup folder first.",
        "warn_same":           "Source folder and backup folder cannot be the same.",
        "confirm_title":       "Confirm",
        "confirm_body":        "{total} file(s) will be processed:\n\n{overview}\n\n-- Step 1: Backup --\nAll originals will be copied to:\n  {backup}\n\n-- Step 2: Convert --\nConverted files will appear in the source folder:\n  {source}\n\n-- Originals --\n{delete_line}\n\nContinue?",
        "delete_yes":          "Originals will be deleted after successful conversion.\nThey are safely stored in the backup folder.",
        "delete_no":           "Original files will remain in the source folder.",
        "done_title":          "Done",
        "done_warn_title":     "Done with warnings",
        "done_body":           "Conversion complete\n\nSuccessful: {success} of {total}\nFailed: {failed}",
        "done_failed":         "\n\nFailed files:\n",
        "done_delete_failed":  "\n\nConverted but could not delete original:\n",
        "done_more":           "... and {n} more",
        "status_done":         "Done — {success} converted, {failed} failed.",
        "status_converting":   "[{n}/{total}] {name}",
        "backup_failed":       "Backup failed: {err}",
        "convert_failed":      "Conversion failed: {err}",
        "delete_failed":       "Delete failed: {err}",
        "done_btn":            "Done",
        "language_btn":        "Language / Taal",
    },
    "nl": {
        "app_title":           "Oud Office Bestandenconverter",
        "app_subtitle":        "Converteert .doc, .ppt en .pps naar moderne formaten met volledig opmaakbehoud via OnlyOffice.",
        "x2t_found":           "OnlyOffice x2t gevonden: {path}",
        "x2t_not_found":       "OnlyOffice x2t NIET gevonden. Plaats de converter/ map naast dit script.",
        "source_folder":       "Bronmap:",
        "backup_folder":       "Backupmap:",
        "x2t_path":            "x2t pad:",
        "choose":              "Kiezen",
        "scan":                "Scannen",
        "scanning":            "Scannen...",
        "convert_btn":         "Backup maken en converteren",
        "converting":          "Bezig...",
        "delete_cb":           "Verwijder originelen na succesvolle conversie",
        "delete_warning":      "  De originelen staan veilig in de backupmap.",
        "not_chosen":          "Nog niet gekozen",
        "x2t_not_found_short": "Niet gevonden — kies handmatig",
        "list_label":          "Gevonden bestanden  (.doc -> .docx   |   .ppt / .pps -> .pptx)",
        "col_name":            "Bestandsnaam",
        "col_ext":             "Extensie",
        "col_converts_to":     "Converteert naar",
        "col_size":            "Grootte",
        "no_files_found":      "Geen ondersteunde bestanden gevonden.",
        "files_found":         "{n} bestanden gevonden:  {summary}",
        "warn_source":         "Kies eerst een bronmap.",
        "warn_x2t":            "OnlyOffice x2t niet gevonden.\nGeef het pad op via de Kiezen knop naast x2t pad.",
        "warn_backup":         "Kies eerst een backupmap.",
        "warn_same":           "Bronmap en backupmap mogen niet hetzelfde zijn.",
        "confirm_title":       "Bevestigen",
        "confirm_body":        "Er worden {total} bestand(en) verwerkt:\n\n{overview}\n\n-- Stap 1: Backup --\nAlle originelen worden gekopieerd naar:\n  {backup}\n\n-- Stap 2: Conversie --\nGeconverteerde bestanden komen in de bronmap:\n  {source}\n\n-- Originelen --\n{delete_line}\n\nDoorgaan?",
        "delete_yes":          "Originelen worden na succesvolle conversie verwijderd uit de bronmap.\nZe staan veilig bewaard in de backupmap.",
        "delete_no":           "Originele bestanden blijven staan in de bronmap.",
        "done_title":          "Klaar",
        "done_warn_title":     "Klaar met waarschuwingen",
        "done_body":           "Conversie voltooid\n\nSuccesvol: {success} van {total}\nMislukt: {failed}",
        "done_failed":         "\n\nMislukte bestanden:\n",
        "done_delete_failed":  "\n\nGeconverteerd maar origineel kon niet worden verwijderd:\n",
        "done_more":           "... en {n} meer",
        "status_done":         "Klaar — {success} geconverteerd, {failed} mislukt.",
        "status_converting":   "[{n}/{total}] {name}",
        "backup_failed":       "Backup mislukt: {err}",
        "convert_failed":      "Conversie mislukt: {err}",
        "delete_failed":       "Verwijderen mislukt: {err}",
        "done_btn":            "Klaar",
        "language_btn":        "Language / Taal",
    }
}


def t(key: str, lang: str, **kwargs) -> str:
    """Return the translated string for key in the given language, formatted with kwargs."""
    text = TRANSLATIONS.get(lang, TRANSLATIONS["en"]).get(key, key)
    return text.format(**kwargs) if kwargs else text


# ---------------------------------------------------------------------------
# Conversion helpers
# ---------------------------------------------------------------------------

def find_x2t() -> Path | None:
    """Search for x2t.exe in the local converter/ folder and standard OnlyOffice install paths."""
    script_dir = Path(__file__).parent
    prog   = os.environ.get("PROGRAMFILES",       r"C:\Program Files")
    prog86 = os.environ.get("PROGRAMFILES(X86)",  r"C:\Program Files (x86)")
    candidates = [
        script_dir / "converter" / "x2t.exe",
        script_dir / "x2t.exe",
        Path(prog)   / "ONLYOFFICE" / "DesktopEditors" / "converter" / "x2t.exe",
        Path(prog86) / "ONLYOFFICE" / "DesktopEditors" / "converter" / "x2t.exe",
    ]
    for path in candidates:
        if path.exists():
            return path
    return None


def convert_with_x2t(x2t: Path, source: Path, dest: Path) -> tuple[bool, str]:
    """Run x2t to convert source to dest. Returns (success, error_message)."""
    try:
        dest.parent.mkdir(parents=True, exist_ok=True)
        kwargs = {}
        if sys.platform == "win32":
            kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW
        result = subprocess.run(
            [str(x2t), str(source), str(dest)],
            capture_output=True, timeout=60, **kwargs
        )
        if result.returncode == 0 and dest.exists():
            return True, ""
        err = (result.stderr.decode('utf-8', errors='ignore').strip() or
               result.stdout.decode('utf-8', errors='ignore').strip())
        return False, f"x2t returncode {result.returncode}: {err or 'no error message'}"
    except subprocess.TimeoutExpired:
        return False, "Conversion timed out (>60s)."
    except Exception as e:
        return False, str(e)


def make_backup(source: Path, source_root: Path, backup_root: Path) -> tuple[bool, str]:
    """Copy source to backup_root, preserving the subfolder structure and file timestamps."""
    try:
        dest = backup_root / source.relative_to(source_root)
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, dest)
        return True, ""
    except Exception as e:
        return False, str(e)


def scan_files(folder: Path) -> list[dict]:
    """Recursively scan folder for supported files. Case-insensitive on all platforms."""
    seen = set()
    results = []
    for path in sorted(folder.rglob('*')):
        ext = path.suffix.lower()
        if ext in SUPPORTED_EXTENSIONS and path not in seen:
            seen.add(path)
            results.append({
                'path':       path,
                'ext':        ext,
                'size':       path.stat().st_size,
                'target_ext': EXTENSION_MAP[ext],
            })
    results.sort(key=lambda x: str(x['path']))
    return results


# ---------------------------------------------------------------------------
# Language selection screen
# ---------------------------------------------------------------------------

class LanguageScreen(ctk.CTkToplevel):
    """Modal window shown on first run to let the user choose a language."""

    def __init__(self, parent, on_chosen):
        super().__init__(parent)
        self.on_chosen = on_chosen
        self.title("Language / Taal")
        self.geometry("400x260")
        self.resizable(False, False)
        self.grab_set()
        self.lift()

        ctk.CTkLabel(self, text="Welcome / Welkom",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(36, 8))
        ctk.CTkLabel(self, text="Please choose your language.\nKies jouw taal.",
                     font=ctk.CTkFont(size=14), text_color=SOFT).pack(pady=(0, 28))

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack()
        ctk.CTkButton(btn_frame, text="English",    width=140, height=44,
                      command=lambda: self._choose("en")).pack(side="left", padx=12)
        ctk.CTkButton(btn_frame, text="Nederlands", width=140, height=44,
                      command=lambda: self._choose("nl")).pack(side="left", padx=12)

    def _choose(self, lang: str):
        self.grab_release()
        self.destroy()
        self.on_chosen(lang)


# ---------------------------------------------------------------------------
# Main application window
# ---------------------------------------------------------------------------

class ConverterApp(ctk.CTk):
    """Main application window. Handles folder selection, scanning and conversion."""

    def __init__(self):
        super().__init__()
        self.geometry("860x700")
        self.resizable(True, True)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        self.source_folder:   Path | None = None
        self.backup_folder:   Path | None = None
        self.found_files:     list[dict]  = []
        self.x2t_path:        Path | None = find_x2t()
        self.delete_originals = ctk.BooleanVar(value=False)

        settings = load_settings()
        self.lang = settings.get("language", None)

        if settings.get("source_folder") and Path(settings["source_folder"]).exists():
            self.source_folder = Path(settings["source_folder"])
        if settings.get("backup_folder") and Path(settings["backup_folder"]).exists():
            self.backup_folder = Path(settings["backup_folder"])
        if settings.get("x2t_path") and Path(settings["x2t_path"]).exists():
            self.x2t_path = Path(settings["x2t_path"])
        if settings.get("delete_originals"):
            self.delete_originals = ctk.BooleanVar(value=settings["delete_originals"])

        if self.lang:
            self._init_ui(self.lang)
        else:
            self.title("Office Legacy Converter")
            self.withdraw()
            self.after(100, self._show_language_screen)

    def _show_language_screen(self):
        self.deiconify()
        LanguageScreen(self, self._on_language_chosen)

    def _on_language_chosen(self, lang: str):
        self.lang = lang
        settings = load_settings()
        settings["language"] = lang
        save_settings(settings)
        self._init_ui(lang)

    def _init_ui(self, lang: str):
        self.title(t("app_title", lang))
        self._build_ui()
        self._refresh_path_labels()

    def _build_ui(self):
        lang = self.lang

        ctk.CTkLabel(self, text=t("app_title", lang),
                     font=ctk.CTkFont(size=22, weight="bold")
                     ).grid(row=0, column=0, padx=24, pady=(20, 2), sticky="w")

        ctk.CTkLabel(self, text=t("app_subtitle", lang),
                     text_color=SOFT, font=ctk.CTkFont(size=13)
                     ).grid(row=1, column=0, padx=24, pady=(0, 4), sticky="w")

        x2t_text  = t("x2t_found",     lang, path=self.x2t_path) if self.x2t_path else t("x2t_not_found", lang)
        x2t_color = ("#1a6e2e", "#4caf7d") if self.x2t_path else ("#8b0000", "#ff6666")
        ctk.CTkLabel(self, text=x2t_text, text_color=x2t_color,
                     font=ctk.CTkFont(size=12), anchor="w"
                     ).grid(row=2, column=0, padx=24, pady=(0, 14), sticky="w")

        map_frame = ctk.CTkFrame(self, fg_color="transparent")
        map_frame.grid(row=3, column=0, padx=24, pady=(0, 10), sticky="ew")
        map_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(map_frame, text=t("source_folder", lang), width=110, anchor="w").grid(row=0, column=0, pady=6, sticky="w")
        self.source_lbl = ctk.CTkLabel(map_frame, text=t("not_chosen", lang), text_color=SOFT, anchor="w")
        self.source_lbl.grid(row=0, column=1, padx=8, sticky="ew")
        ctk.CTkButton(map_frame, text=t("choose", lang), width=90, command=self._pick_source).grid(row=0, column=2)

        ctk.CTkLabel(map_frame, text=t("backup_folder", lang), width=110, anchor="w").grid(row=1, column=0, pady=6, sticky="w")
        self.backup_lbl = ctk.CTkLabel(map_frame, text=t("not_chosen", lang), text_color=SOFT, anchor="w")
        self.backup_lbl.grid(row=1, column=1, padx=8, sticky="ew")
        ctk.CTkButton(map_frame, text=t("choose", lang), width=90, command=self._pick_backup).grid(row=1, column=2)

        ctk.CTkLabel(map_frame, text=t("x2t_path", lang), width=110, anchor="w").grid(row=2, column=0, pady=6, sticky="w")
        self.x2t_lbl = ctk.CTkLabel(
            map_frame,
            text=str(self.x2t_path) if self.x2t_path else t("x2t_not_found_short", lang),
            text_color=SOFT, anchor="w", font=ctk.CTkFont(size=11)
        )
        self.x2t_lbl.grid(row=2, column=1, padx=8, sticky="ew")
        ctk.CTkButton(map_frame, text=t("choose", lang), width=90, command=self._pick_x2t).grid(row=2, column=2)

        top_btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_btn_frame.grid(row=4, column=0, padx=24, pady=(0, 6), sticky="w")
        self.btn_scan = ctk.CTkButton(top_btn_frame, text=t("scan", lang), command=self._scan, height=36)
        self.btn_scan.pack(side="left", padx=(0, 12))
        self.btn_lang = ctk.CTkButton(top_btn_frame, text=t("language_btn", lang),
                                      fg_color="gray", hover_color="#555555", height=36,
                                      command=self._change_language)
        self.btn_lang.pack(side="left")

        delete_frame = ctk.CTkFrame(self, fg_color="transparent")
        delete_frame.grid(row=5, column=0, padx=24, pady=(0, 8), sticky="w")
        self.delete_cb = ctk.CTkCheckBox(delete_frame, text=t("delete_cb", lang),
                                         variable=self.delete_originals,
                                         command=self._on_delete_toggle,
                                         font=ctk.CTkFont(size=13))
        self.delete_cb.pack(side="left")
        self.delete_warning = ctk.CTkLabel(delete_frame, text=t("delete_warning", lang),
                                           text_color=("#8a5a00", "#d4a017"),
                                           font=ctk.CTkFont(size=12))

        self.list_frame = ctk.CTkScrollableFrame(self, fg_color="transparent",
                                                 label_text=t("list_label", lang))
        self.list_frame.grid(row=6, column=0, padx=24, pady=(0, 10), sticky="nsew")
        self.grid_rowconfigure(6, weight=1)

        self.status_lbl = ctk.CTkLabel(self, text="", text_color=SOFT,
                                       font=ctk.CTkFont(size=12), anchor="w")
        self.status_lbl.grid(row=7, column=0, padx=24, pady=(0, 4), sticky="ew")

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=8, column=0, padx=24, pady=(0, 20), sticky="ew")
        self.btn_convert = ctk.CTkButton(btn_frame, text=t("convert_btn", lang),
                                         command=self._start_conversion, height=42, state="disabled")
        self.btn_convert.pack(side="left")
        self.progress = ctk.CTkProgressBar(btn_frame, width=220)
        self.progress.pack(side="left", padx=(16, 0))
        self.progress.set(0)
        self.progress_lbl = ctk.CTkLabel(btn_frame, text="", text_color=SOFT, font=ctk.CTkFont(size=12))
        self.progress_lbl.pack(side="left", padx=8)

        if self.delete_originals.get():
            self.delete_warning.pack(side="left")

    def _change_language(self):
        LanguageScreen(self, self._rebuild_language)

    def _rebuild_language(self, lang: str):
        """Rebuild the entire UI in the newly chosen language."""
        self.lang = lang
        settings = load_settings()
        settings["language"] = lang
        save_settings(settings)
        for widget in self.winfo_children():
            widget.destroy()
        self.title(t("app_title", lang))
        self._build_ui()
        self._refresh_path_labels()
        if self.found_files:
            self._show_results(self.found_files)

    def _refresh_path_labels(self):
        """Populate path labels with values restored from settings."""
        if self.source_folder:
            s = str(self.source_folder)
            self.source_lbl.configure(text=f"...{s[-50:]}" if len(s) > 52 else s)
        if self.backup_folder:
            s = str(self.backup_folder)
            self.backup_lbl.configure(text=f"...{s[-50:]}" if len(s) > 52 else s)
        if self.x2t_path:
            self.x2t_lbl.configure(text=str(self.x2t_path))
        if self.delete_originals.get():
            self.delete_warning.pack(side="left")

    def _save_settings(self):
        settings = load_settings()
        settings.update({
            "language":        self.lang,
            "source_folder":   str(self.source_folder)   if self.source_folder   else None,
            "backup_folder":   str(self.backup_folder)   if self.backup_folder   else None,
            "x2t_path":        str(self.x2t_path)        if self.x2t_path        else None,
            "delete_originals": self.delete_originals.get(),
        })
        save_settings(settings)

    def _on_delete_toggle(self):
        if self.delete_originals.get():
            self.delete_warning.pack(side="left")
        else:
            self.delete_warning.pack_forget()
        self._save_settings()

    def _pick_source(self):
        folder = filedialog.askdirectory()
        if folder:
            self.source_folder = Path(folder)
            s = folder
            self.source_lbl.configure(text=f"...{s[-50:]}" if len(s) > 52 else s)
            self._clear_list()
            self.btn_convert.configure(state="disabled")
            self._save_settings()

    def _pick_backup(self):
        folder = filedialog.askdirectory()
        if folder:
            self.backup_folder = Path(folder)
            s = folder
            self.backup_lbl.configure(text=f"...{s[-50:]}" if len(s) > 52 else s)
            self._save_settings()

    def _pick_x2t(self):
        path = filedialog.askopenfilename(filetypes=[("x2t converter", "x2t.exe"), ("All files", "*.*")])
        if path:
            self.x2t_path = Path(path)
            self.x2t_lbl.configure(text=str(self.x2t_path))
            self._save_settings()

    def _clear_list(self):
        for w in self.list_frame.winfo_children():
            w.destroy()
        self.found_files = []

    def _set_controls_enabled(self, enabled: bool):
        """Enable or disable all interactive controls during a background operation."""
        state = "normal" if enabled else "disabled"
        for widget in self.winfo_children():
            if isinstance(widget, ctk.CTkFrame):
                for child in widget.winfo_children():
                    if isinstance(child, ctk.CTkButton):
                        child.configure(state=state)
        self.btn_scan.configure(state=state)
        self.btn_lang.configure(state=state)
        self.btn_convert.configure(state=state)

    def _scan(self):
        lang = self.lang
        if not self.source_folder:
            messagebox.showwarning("!", t("warn_source", lang))
            return
        if not self.x2t_path:
            messagebox.showerror("!", t("warn_x2t", lang))
            return
        self._clear_list()
        self.btn_scan.configure(state="disabled", text=t("scanning", lang))
        self.btn_lang.configure(state="disabled")
        self.status_lbl.configure(text=t("scanning", lang))

        def do_scan():
            files = scan_files(self.source_folder)
            def _done():
                self._show_results(files)
                self.btn_lang.configure(state="normal")
            self.after(0, _done)

        threading.Thread(target=do_scan, daemon=True).start()

    def _show_results(self, files: list[dict]):
        lang = self.lang
        self.found_files = files
        self.btn_scan.configure(state="normal", text=t("scan", lang))

        if not files:
            self.status_lbl.configure(text=t("no_files_found", lang))
            return

        counts = {}
        for f in files:
            counts[f['ext']] = counts.get(f['ext'], 0) + 1
        summary = "  |  ".join(f"{v}x {k}" for k, v in sorted(counts.items()))
        self.status_lbl.configure(text=t("files_found", lang, n=len(files), summary=summary))

        header = ctk.CTkFrame(self.list_frame, fg_color="transparent")
        header.pack(fill="x", pady=(0, 2))
        for col, (key, width, anchor) in enumerate([
            ("col_name", 300, "w"), ("col_ext", 80, "w"),
            ("col_converts_to", 120, "w"), ("col_size", 80, "e"),
        ]):
            ctk.CTkLabel(header, text=t(key, lang), font=ctk.CTkFont(weight="bold"),
                         anchor=anchor, width=width).grid(row=0, column=col, sticky=anchor)

        ctk.CTkFrame(self.list_frame, height=1, fg_color="gray").pack(fill="x", pady=(0, 4))

        ext_colors = {
            '.doc': ("#1a4f8a", "#5b9bd5"),
            '.ppt': ("#8a3a1a", "#d5825b"),
            '.pps': ("#6a3a8a", "#b57bd5"),
        }

        for f in files:
            row = ctk.CTkFrame(self.list_frame, fg_color="transparent")
            row.pack(fill="x", pady=1)
            name    = f['path'].name
            display = f"...{name[-38:]}" if len(name) > 40 else name
            size_kb = f"{max(1, f['size'] // 1024)} KB"
            color   = ext_colors.get(f['ext'], SOFT)
            for col, (text, width, anchor, tc) in enumerate([
                (display,          300, "w", None),
                (f['ext'],          80, "w", color),
                (f['target_ext'],  120, "w", SOFT),
                (size_kb,           80, "e", SOFT),
            ]):
                kwargs = {"text_color": tc} if tc else {}
                ctk.CTkLabel(row, text=text, anchor=anchor, width=width,
                             font=ctk.CTkFont(size=12), **kwargs).grid(row=0, column=col, sticky=anchor)

        self.btn_convert.configure(state="normal")

    def _start_conversion(self):
        lang = self.lang
        if not self.backup_folder:
            messagebox.showwarning("!", t("warn_backup", lang))
            return
        if self.source_folder.resolve() == self.backup_folder.resolve():
            messagebox.showwarning("!", t("warn_same", lang))
            return

        total    = len(self.found_files)
        counts   = {}
        for f in self.found_files:
            counts[f['ext']] = counts.get(f['ext'], 0) + 1
        overview    = "\n".join(f"  {v}x {k} -> {EXTENSION_MAP[k]}" for k, v in sorted(counts.items()))
        delete_line = t("delete_yes", lang) if self.delete_originals.get() else t("delete_no", lang)

        if not messagebox.askyesno(
            t("confirm_title", lang),
            t("confirm_body", lang, total=total, overview=overview,
              backup=self.backup_folder, source=self.source_folder, delete_line=delete_line)
        ):
            return

        self._set_controls_enabled(False)
        self.btn_convert.configure(text=t("converting", lang))
        self.progress.set(0)
        threading.Thread(target=self._run_conversion, daemon=True).start()

    def _run_conversion(self):
        """Worker thread: backs up and converts each file, then calls _show_report on the main thread."""
        lang         = self.lang
        total        = len(self.found_files)
        success      = 0
        failed       = []
        delete_failed = []

        for i, f in enumerate(self.found_files):
            source = f['path']
            n      = i + 1
            self.after(0, lambda nm=source.name, idx=n: (
                self.status_lbl.configure(text=t("status_converting", lang, n=idx, total=total, name=nm)),
                self.progress.set(idx / total),
                self.progress_lbl.configure(text=f"{idx}/{total}")
            ))

            dest   = source.with_suffix(f['target_ext'])
            ok, err = make_backup(source, self.source_folder, self.backup_folder)
            if not ok:
                failed.append((source.name, t("backup_failed", lang, err=err)))
                continue

            ok, err = convert_with_x2t(self.x2t_path, source, dest)
            if ok:
                success += 1
                try:
                    stat = source.stat()
                    os.utime(dest, (stat.st_atime, stat.st_mtime))
                except Exception:
                    pass
                if self.delete_originals.get():
                    try:
                        source.unlink()
                    except Exception as e:
                        delete_failed.append((source.name, t("delete_failed", lang, err=e)))
            else:
                failed.append((source.name, t("convert_failed", lang, err=err)))

            time.sleep(0.02)

        self.after(0, lambda: self._show_report(total, success, failed, delete_failed))

    def _show_report(self, total, success, failed, delete_failed=None):
        """Re-enable controls and display the conversion summary."""
        lang = self.lang
        self._set_controls_enabled(True)
        self.btn_convert.configure(text=t("convert_btn", lang))
        if not self.found_files:
            self.btn_convert.configure(state="disabled")
        self.progress.set(1)
        self.progress_lbl.configure(text=t("done_btn", lang))
        self.status_lbl.configure(text=t("status_done", lang, success=success, failed=len(failed)))

        delete_failed = delete_failed or []
        rapport = t("done_body", lang, success=success, total=total, failed=len(failed))
        if delete_failed:
            rapport += t("done_delete_failed", lang)
            rapport += "\n".join(f"  {name}: {err}" for name, err in delete_failed[:6])
        if failed:
            rapport += t("done_failed", lang)
            rapport += "\n".join(f"  {name}: {err}" for name, err in failed[:12])
            if len(failed) > 12:
                rapport += "\n" + t("done_more", lang, n=len(failed) - 12)

        title = t("done_title", lang) if not failed else t("done_warn_title", lang)
        if not failed:
            messagebox.showinfo(title, rapport)
        else:
            messagebox.showwarning(title, rapport)


if __name__ == "__main__":
    app = ConverterApp()
    app.mainloop()
