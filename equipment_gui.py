#!/usr/bin/env python3
"""
Equipment Counter — GUI (tkinter).

Provides a graphical interface for batch processing engineering drawings:
  - Folder selection
  - Options: convert DWG, keep converted, parse DXF/PDF
  - Output JSON filename
  - Real-time log output
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

from batch_equipment import run_pipeline, find_oda, scan_files

# ---------------------------------------------------------------------------
#  Constants
# ---------------------------------------------------------------------------

APP_TITLE = "Подсчёт оборудования — Equipment Counter"
DEFAULT_JSON = "equipment_report.json"
WIN_WIDTH, WIN_HEIGHT = 780, 620
BG = "#f5f5f5"
ACCENT = "#1a73e8"
FONT = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
FONT_LOG = ("Consolas", 9)


class EquipmentApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}")
        self.root.minsize(600, 480)
        self.root.configure(bg=BG)

        self._running = False

        # ── Variables ──
        self.var_folder = tk.StringVar()
        self.var_convert_dwg = tk.BooleanVar(value=True)
        self.var_keep_converted = tk.BooleanVar(value=True)
        self.var_parse_dxf = tk.BooleanVar(value=True)
        self.var_parse_pdf = tk.BooleanVar(value=False)
        self.var_json_name = tk.StringVar(value=DEFAULT_JSON)
        self.var_generate_png = tk.BooleanVar(value=True)
        self.var_png_dpi = tk.IntVar(value=150)
        self.var_generate_vor = tk.BooleanVar(value=True)
        self.var_combined_vor = tk.BooleanVar(value=True)

        self._build_ui()
        self._check_oda()

    # ------------------------------------------------------------------
    #  UI Construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        pad = dict(padx=12, pady=4)

        # ── Title ──
        title_frame = tk.Frame(self.root, bg=ACCENT, height=44)
        title_frame.pack(fill="x")
        title_frame.pack_propagate(False)
        tk.Label(
            title_frame, text="⚡  " + APP_TITLE,
            bg=ACCENT, fg="white", font=("Segoe UI", 12, "bold"),
        ).pack(side="left", padx=16)

        main = tk.Frame(self.root, bg=BG)
        main.pack(fill="both", expand=True, padx=12, pady=8)

        # ── Folder ──
        folder_frame = tk.LabelFrame(
            main, text="  Папка с чертежами  ", font=FONT_BOLD, bg=BG, padx=8, pady=6,
        )
        folder_frame.pack(fill="x", **pad)

        entry_row = tk.Frame(folder_frame, bg=BG)
        entry_row.pack(fill="x")

        self.folder_entry = tk.Entry(
            entry_row, textvariable=self.var_folder, font=FONT,
        )
        self.folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.btn_browse = tk.Button(
            entry_row, text="Обзор…", font=FONT, command=self._browse_folder,
            bg="white", relief="groove", cursor="hand2",
        )
        self.btn_browse.pack(side="right")

        self.lbl_scan = tk.Label(
            folder_frame, text="", font=("Segoe UI", 9), bg=BG, fg="#666",
        )
        self.lbl_scan.pack(anchor="w", pady=(4, 0))

        self.var_folder.trace_add("write", lambda *_: self._on_folder_change())

        # ── Options ──
        opts_frame = tk.LabelFrame(
            main, text="  Настройки  ", font=FONT_BOLD, bg=BG, padx=8, pady=6,
        )
        opts_frame.pack(fill="x", **pad)

        row1 = tk.Frame(opts_frame, bg=BG)
        row1.pack(fill="x")
        self.chk_convert = tk.Checkbutton(
            row1, text="Конвертировать DWG → DXF",
            variable=self.var_convert_dwg, font=FONT, bg=BG,
            command=self._on_convert_toggle,
        )
        self.chk_convert.pack(side="left")
        self.chk_keep = tk.Checkbutton(
            row1, text="Не удалять сконвертированные DXF",
            variable=self.var_keep_converted, font=FONT, bg=BG,
        )
        self.chk_keep.pack(side="left", padx=(24, 0))

        row2 = tk.Frame(opts_frame, bg=BG)
        row2.pack(fill="x", pady=(4, 0))
        tk.Checkbutton(
            row2, text="Парсить DXF",
            variable=self.var_parse_dxf, font=FONT, bg=BG,
        ).pack(side="left")
        tk.Checkbutton(
            row2, text="Парсить PDF",
            variable=self.var_parse_pdf, font=FONT, bg=BG,
        ).pack(side="left", padx=(24, 0))

        row3 = tk.Frame(opts_frame, bg=BG)
        row3.pack(fill="x", pady=(4, 0))
        tk.Checkbutton(
            row3, text="Генерировать PNG с маркерами оборудования",
            variable=self.var_generate_png, font=FONT, bg=BG,
        ).pack(side="left")
        tk.Label(row3, text="DPI:", font=FONT, bg=BG).pack(side="left", padx=(16, 0))
        dpi_spin = tk.Spinbox(
            row3, from_=72, to=300, increment=50,
            textvariable=self.var_png_dpi, font=FONT, width=5,
        )

        row4 = tk.Frame(opts_frame, bg=BG)
        row4.pack(fill="x", pady=(4, 0))
        tk.Checkbutton(
            row4, text="Генерировать ВОР (.docx)",
            variable=self.var_generate_vor, font=FONT, bg=BG,
        ).pack(side="left")
        tk.Checkbutton(
            row4, text="Объединённый ВОР (ЭО+ЭМ+ЭГ)",
            variable=self.var_combined_vor, font=FONT, bg=BG,
        ).pack(side="left", padx=(16, 0))
        dpi_spin.pack(side="left", padx=(4, 0))

        # ── JSON name ──
        json_frame = tk.Frame(opts_frame, bg=BG)
        json_frame.pack(fill="x", pady=(8, 0))
        tk.Label(
            json_frame, text="Имя JSON-отчёта:", font=FONT, bg=BG,
        ).pack(side="left")
        tk.Entry(
            json_frame, textvariable=self.var_json_name, font=FONT, width=35,
        ).pack(side="left", padx=(8, 0))

        # ── ODA status ──
        self.lbl_oda = tk.Label(
            opts_frame, text="", font=("Segoe UI", 9), bg=BG,
        )
        self.lbl_oda.pack(anchor="w", pady=(6, 0))

        # ── Run button ──
        btn_frame = tk.Frame(main, bg=BG)
        btn_frame.pack(fill="x", **pad)
        self.btn_run = tk.Button(
            btn_frame, text="▶  Запуск", font=("Segoe UI", 11, "bold"),
            bg=ACCENT, fg="white", activebackground="#1557b0",
            activeforeground="white", relief="flat", cursor="hand2",
            command=self._start, padx=24, pady=6,
        )
        self.btn_run.pack(side="left")

        self.progress = ttk.Progressbar(btn_frame, mode="indeterminate", length=200)
        self.progress.pack(side="left", padx=(16, 0))

        # ── Log ──
        log_frame = tk.LabelFrame(
            main, text="  Лог  ", font=FONT_BOLD, bg=BG, padx=4, pady=4,
        )
        log_frame.pack(fill="both", expand=True, **pad)

        self.log_text = tk.Text(
            log_frame, font=FONT_LOG, bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white", wrap="word", state="disabled",
            relief="flat", borderwidth=0,
        )
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

    # ------------------------------------------------------------------
    #  Helpers
    # ------------------------------------------------------------------

    def _check_oda(self):
        oda = find_oda()
        if oda:
            self.lbl_oda.config(
                text=f"✓ ODA File Converter: {Path(oda).name}", fg="#2e7d32",
            )
        else:
            self.lbl_oda.config(
                text="⚠ ODA File Converter не найден (конвертация DWG недоступна)", fg="#c62828",
            )
            self.var_convert_dwg.set(False)
            self.chk_convert.config(state="disabled")
            self.chk_keep.config(state="disabled")

    def _on_convert_toggle(self):
        state = "normal" if self.var_convert_dwg.get() else "disabled"
        self.chk_keep.config(state=state)

    def _browse_folder(self):
        d = filedialog.askdirectory(title="Выберите папку с чертежами")
        if d:
            self.var_folder.set(d)

    def _on_folder_change(self):
        folder = self.var_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            self.lbl_scan.config(text="")
            return
        try:
            dwg, dxf, pdf = scan_files(Path(folder))
            self.lbl_scan.config(
                text=f"Найдено:  DWG: {len(dwg)}    DXF: {len(dxf)}    PDF: {len(pdf)}",
            )
        except Exception:
            self.lbl_scan.config(text="")

    def log(self, text: str):
        """Thread-safe logging to the text widget."""
        self.root.after(0, self._append_log, text)

    def _append_log(self, text: str):
        self.log_text.config(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    # ------------------------------------------------------------------
    #  Run pipeline
    # ------------------------------------------------------------------

    def _start(self):
        folder = self.var_folder.get().strip()
        if not folder:
            messagebox.showwarning("Ошибка", "Укажите папку с чертежами")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("Ошибка", f"Папка не найдена:\n{folder}")
            return

        json_name = self.var_json_name.get().strip()
        if not json_name:
            json_name = DEFAULT_JSON
        if not json_name.endswith(".json"):
            json_name += ".json"

        self._set_running(True)
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

        t = threading.Thread(
            target=self._run_worker,
            args=(Path(folder), json_name),
            daemon=True,
        )
        t.start()

    def _run_worker(self, folder: Path, json_name: str):
        try:
            run_pipeline(
                folder,
                convert_dwg=self.var_convert_dwg.get(),
                keep_converted=self.var_keep_converted.get(),
                parse_dxf=self.var_parse_dxf.get(),
                parse_pdf=self.var_parse_pdf.get(),
                generate_png=self.var_generate_png.get(),
                png_dpi=self.var_png_dpi.get(),
                output_json=json_name,
                log=self.log,
            )
            if self.var_generate_vor.get():
                self.log("\n--- Генерация ВОР ---")
                try:
                    from vor_generator import generate_vor
                    generate_vor(folder, log=self.log)
                except ImportError as exc:
                    self.log(f"⚠ ВОР: не установлен python-docx: {exc}")
                except Exception as exc:
                    self.log(f"⚠ ВОР ошибка: {exc}")

            if self.var_combined_vor.get():
                try:
                    from vor_generator import generate_vor_combined, find_dwg_parent
                    dwg_parent = find_dwg_parent(folder)
                    if dwg_parent:
                        self.log("\n--- Генерация объединённого ВОР ---")
                        result = generate_vor_combined(dwg_parent, log=self.log)
                        self.log(f"Объединённый ВОР: {result}")
                    else:
                        self.log("\n  Объединённый ВОР: не найдены разделы-соседи (ЭО/ЭМ/ЭГ)")
                except ImportError as exc:
                    self.log(f"⚠ Объединённый ВОР: {exc}")
                except Exception as exc:
                    self.log(f"⚠ Объединённый ВОР ошибка: {exc}")
        except Exception as e:
            self.log(f"\n⚠ ОШИБКА: {e}")
        finally:
            self.root.after(0, self._set_running, False)

    def _set_running(self, running: bool):
        self._running = running
        if running:
            self.btn_run.config(state="disabled", text="⏳ Обработка…")
            self.btn_browse.config(state="disabled")
            self.progress.start(15)
        else:
            self.btn_run.config(state="normal", text="▶  Запуск")
            self.btn_browse.config(state="normal")
            self.progress.stop()


def main():
    root = tk.Tk()
    try:
        root.iconbitmap(default="")
    except Exception:
        pass
    EquipmentApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
