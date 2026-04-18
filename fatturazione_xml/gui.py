"""
gui.py

Tkinter GUI for the FatturaPA XML export tool.

All tkinter calls are guarded inside run() so the module can be imported
without starting the GUI event loop (important for tests).
"""

from __future__ import annotations

import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .config import (
    get_output_path,
    is_configured,
    load_config,
    save_config,
)
from .excel_reader import read_cell_values
from .xlsm_parser import get_sheet_bindings
from .xml_builder import build_xml

EXPORT_TYPES = ["XML con IVA", "XML intento", "XML no IVA", "XML Estere"]

_HEADER_BG    = "#1B2838"
_EXPORT_BTN   = "#2563EB"
_EXPORT_BTN_H = "#1D4ED8"
_FOOTER_BG    = "#F0F0F0"
_STATUS_IDLE  = "#6B7280"
_STATUS_OK    = "#16A34A"
_STATUS_ERR   = "#DC2626"
_LABEL_MUTED  = "#555555"

# Safe font that is available on macOS in tkinter
_FONT_FAMILY = "Helvetica Neue"


def _font(size: int, weight: str = "normal") -> tuple:
    return (_FONT_FAMILY, size, weight)


class _App(tk.Tk):
    def __init__(self, initial_type: str | None = None) -> None:
        super().__init__()

        self.title("Fatturazione XML")
        self.geometry("440x250")
        self.resizable(False, False)
        self.configure(bg="white")

        self._initial_type = initial_type
        self._build_ui()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        # ── Header ──────────────────────────────────────────────────────
        header = tk.Frame(self, bg=_HEADER_BG, height=52)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        tk.Label(
            header, text="Fatturazione XML",
            bg=_HEADER_BG, fg="white",
            font=_font(14, "bold"),
        ).place(relx=0.5, rely=0.36, anchor="center")

        tk.Label(
            header, text="FatturaPA v1.2  ·  AdE",
            bg=_HEADER_BG, fg="#8BA3B8",
            font=_font(9),
        ).place(relx=0.5, rely=0.74, anchor="center")

        # ── Footer ──────────────────────────────────────────────────────
        footer = tk.Frame(self, bg=_FOOTER_BG, height=34)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        self._status_var = tk.StringVar(value="Pronto.")
        self._status_lbl = tk.Label(
            footer,
            textvariable=self._status_var,
            bg=_FOOTER_BG, fg=_STATUS_IDLE,
            font=_font(9),
        )
        self._status_lbl.pack(side="left", padx=10)

        settings_lbl = tk.Label(
            footer, text="⚙  Impostazioni",
            bg=_FOOTER_BG, fg=_LABEL_MUTED,
            font=_font(9),
        )
        settings_lbl.pack(side="right", padx=10)
        settings_lbl.bind("<Button-1>", lambda _e: self._open_settings())
        settings_lbl.bind("<Enter>",    lambda _e: settings_lbl.config(fg="#111111"))
        settings_lbl.bind("<Leave>",    lambda _e: settings_lbl.config(fg=_LABEL_MUTED))

        # ── Main content ─────────────────────────────────────────────────
        # A single centred column, 360 px wide, via grid
        content = tk.Frame(self, bg="white")
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=1)
        content.rowconfigure(0, weight=1)

        # Centre wrapper
        wrapper = tk.Frame(content, bg="white")
        wrapper.grid(row=0, column=0)          # centred by the weight above

        # Force wrapper column to exactly 360 px so combo and button align
        wrapper.columnconfigure(0, minsize=360)

        tk.Label(
            wrapper, text="Tipo fattura:",
            bg="white", fg=_LABEL_MUTED,
            font=_font(10),
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        self._export_var = tk.StringVar(
            value=self._initial_type
            if self._initial_type in EXPORT_TYPES
            else EXPORT_TYPES[0]
        )
        # tk.OptionMenu renders as a native macOS popup button — no border, no focus ring
        option_menu = tk.OptionMenu(wrapper, self._export_var, *EXPORT_TYPES)
        option_menu.config(
            font=_font(12),
            relief="flat",
            bd=0,
            highlightthickness=0,
            bg="white",
            activebackground="#E5E7EB",
        )
        option_menu["menu"].config(font=_font(11), bd=0)
        option_menu.grid(row=1, column=0, sticky="ew", ipady=3)

        # Export button — tk.Button ignores bg on macOS; use Label instead
        self._export_btn = tk.Label(
            wrapper,
            text="Esporta XML",
            bg=_EXPORT_BTN, fg="white",
            font=_font(13, "bold"),
            pady=10,
        )
        self._export_btn.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        self._export_btn.bind("<Button-1>", lambda _e: self._do_export())
        self._export_btn.bind("<Enter>",    lambda _e: self._export_btn.config(bg=_EXPORT_BTN_H))
        self._export_btn.bind("<Leave>",    lambda _e: self._export_btn.config(bg=_EXPORT_BTN))

    # ------------------------------------------------------------------
    # Export logic
    # ------------------------------------------------------------------

    def _do_export(self) -> None:
        try:
            config = load_config()

            if not is_configured():
                self._open_settings()
                return

            xlsm_path = config["xlsm_path"]
            sheet_name = self._export_var.get()

            sheet_bindings = get_sheet_bindings(xlsm_path, sheet_name)

            try:
                values = read_cell_values(xlsm_path, sheet_name, sheet_bindings.bindings)
            except ValueError as exc:
                self._set_status(f"✗ {exc}", _STATUS_ERR)
                messagebox.showerror(
                    "Errore",
                    f"Valori non disponibili.\n\nSalva il file Excel e riprova.\n\n{exc}",
                )
                return

            # Extract ProgressivoInvio directly from the already-computed binding values
            progressivo = None
            for binding, value in values:
                if "ProgressivoInvio" in binding.xpath and value is not None:
                    progressivo = str(value).strip()
                    break
            if not progressivo:
                raise ValueError(
                    "ProgressivoInvio non trovato. Salva il file Excel e riprova."
                )

            xml_str = build_xml(values, sheet_bindings.xml_map)
            output_path = get_output_path(config, progressivo)
            output_path.write_text(xml_str, encoding="utf-8")

            self._set_status(f"✓  {output_path.name}", _STATUS_OK, auto_reset=True)

        except Exception as exc:
            print(f"Export error: {exc}", file=sys.stderr)
            self._set_status(f"✗  {str(exc)[:100]}", _STATUS_ERR, auto_reset=True)
            messagebox.showerror("Errore", str(exc))

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _set_status(self, text: str, color: str = _STATUS_IDLE, auto_reset: bool = False) -> None:
        self._status_var.set(text)
        self._status_lbl.config(fg=color)
        if auto_reset:
            self.after(3000, lambda: self._set_status("Pronto."))

    # ------------------------------------------------------------------
    # Settings dialog
    # ------------------------------------------------------------------

    def _open_settings(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Impostazioni")
        dialog.geometry("520x260")
        dialog.resizable(False, False)
        dialog.grab_set()

        config = load_config()

        frame = ttk.LabelFrame(dialog, text="Configurazione", padding=(16, 10, 16, 8))
        frame.pack(fill="both", expand=True, padx=16, pady=(12, 4))
        frame.columnconfigure(1, weight=1)

        entries: dict[str, tk.StringVar] = {}

        def add_row(label: str, key: str, row: int, browse_type: str | None = None) -> None:
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=4)
            var = tk.StringVar(value=str(config.get(key, "")))
            entries[key] = var
            ttk.Entry(frame, textvariable=var, width=42).grid(
                row=row, column=1, sticky="ew", padx=(8, 0), pady=4
            )
            if browse_type == "file":
                def _pick(v=var):
                    p = filedialog.askopenfilename(
                        parent=dialog, title="Seleziona file xlsm",
                        filetypes=[("Excel macro-enabled", "*.xlsm"), ("Tutti", "*.*")],
                    )
                    if p:
                        v.set(p)
                ttk.Button(frame, text="Sfoglia…", command=_pick).grid(
                    row=row, column=2, padx=(4, 0), pady=4
                )
            elif browse_type == "dir":
                def _pick(v=var):
                    p = filedialog.askdirectory(parent=dialog, title="Seleziona cartella")
                    if p:
                        v.set(p)
                ttk.Button(frame, text="Sfoglia…", command=_pick).grid(
                    row=row, column=2, padx=(4, 0), pady=4
                )

        add_row("File xlsm:", "xlsm_path", 0, "file")
        add_row("Cartella output XML:", "xml_output_dir", 1, "dir")
        add_row("Prefisso nome file:", "filename_prefix", 2)
        add_row("Anno:", "year", 3)

        btn_frame = ttk.Frame(dialog, padding=(16, 4, 16, 10))
        btn_frame.pack(fill="x", side="bottom")

        def _save() -> None:
            cfg = {
                "xlsm_path":       entries["xlsm_path"].get().strip(),
                "xml_output_dir":  entries["xml_output_dir"].get().strip(),
                "filename_prefix": entries["filename_prefix"].get().strip(),
                "year":            entries["year"].get().strip(),
            }
            if not Path(cfg["xlsm_path"]).exists():
                messagebox.showerror("Errore", f"File xlsm non trovato:\n{cfg['xlsm_path']}", parent=dialog)
                return
            try:
                cfg["year"] = int(cfg["year"])
            except ValueError:
                messagebox.showerror("Errore", "Anno deve essere un numero intero (es. 2026).", parent=dialog)
                return
            try:
                Path(cfg["xml_output_dir"]).mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Errore", f"Cartella output non accessibile:\n{e}", parent=dialog)
                return
            save_config(cfg)
            dialog.destroy()
            self._set_status("Impostazioni salvate.", _STATUS_IDLE)

        ttk.Button(btn_frame, text="Salva",   command=_save,              width=10).pack(side="right", padx=(4, 0))
        ttk.Button(btn_frame, text="Annulla", command=dialog.destroy,     width=10).pack(side="right")


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def run(initial_type: str | None = None) -> None:
    """Launch the tkinter GUI. Pass initial_type to pre-select the export type."""
    app = _App(initial_type=initial_type)
    app.mainloop()
