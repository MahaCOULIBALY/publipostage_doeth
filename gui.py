#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interface graphique Publipostage DOETH — Groupe Interaction.
Charte graphique : orange #E85D04 / marine #0F2A4A / blanc #FFFFFF.
"""

import argparse
import logging
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from src.utils.config import get
from src.utils.logger import get_logger
from src.document_generator import OutputFormat


# ── Charte Groupe Interaction ──────────────────────────────────────────────────
class _Brand:
    """Constantes de la charte graphique Groupe Interaction."""
    ORANGE = "#E85D04"
    ORANGE_HOVER = "#C94E03"
    NAVY = "#0F2A4A"
    NAVY_LIGHT = "#1A3E6A"
    WHITE = "#FFFFFF"
    GRAY_BG = "#F5F5F5"
    GRAY_BORDER = "#E0E0E0"
    GRAY_TEXT = "#4A4A4A"
    FONT = "Segoe UI"


# ── Handler de log coloré ──────────────────────────────────────────────────────
class _LogHandler(logging.Handler):
    _TAGS = {
        logging.DEBUG: "log_debug",
        logging.INFO: "log_info",
        logging.WARNING: "log_warning",
        logging.ERROR: "log_error",
        logging.CRITICAL: "log_error",
    }

    def __init__(self, widget: tk.Text) -> None:
        super().__init__()
        self._w = widget

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        tag = self._TAGS.get(record.levelno, "log_info")
        self._w.after(0, lambda: self._insert(msg, tag))

    def _insert(self, msg: str, tag: str) -> None:
        self._w.configure(state="normal")
        self._w.insert("end", msg + "\n", tag)
        self._w.see("end")
        self._w.configure(state="disabled")


# ── Application ────────────────────────────────────────────────────────────────
class PublipostageGUI:

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.processing = False

        self._init_vars()
        self._apply_theme()
        self._build_ui()
        self._setup_logging()

        root.title("Publipostage DOETH — Groupe Interaction")
        root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Variables ─────────────────────────────────────────────────────────────
    def _init_vars(self) -> None:
        self.input_file_var = tk.StringVar(value=str(
            Path(get('paths.input_dir', './data/input')) /
            get('defaults.input_filename', 'donnees.xlsx')))
        self.sheet_var = tk.StringVar(value=get('defaults.excel_sheet', 'Feuil1'))
        self.output_dir_var = tk.StringVar(value=get('paths.output_dir', './data/output'))
        self.csv_path_var = tk.StringVar()
        self.logo_var = tk.StringVar(value=get('resources.logo_path', ''))
        self.sig_var = tk.StringVar(value=get('resources.signature_path', ''))
        self.skip_var = tk.BooleanVar(value=False)
        self.debug_var = tk.BooleanVar(value=False)
        self.fmt_var = tk.StringVar(value="docx")

    # ── Theme ttk ─────────────────────────────────────────────────────────────
    def _apply_theme(self) -> None:
        B = _Brand
        s = ttk.Style()
        if 'clam' in s.theme_names():
            s.theme_use('clam')

        s.configure('TFrame', background=B.GRAY_BG)
        s.configure('TLabel', background=B.GRAY_BG, foreground=B.GRAY_TEXT,
                    font=(B.FONT, 10))
        s.configure('TLabelframe', background=B.GRAY_BG)
        s.configure('TLabelframe.Label', background=B.GRAY_BG,
                    foreground=B.NAVY, font=(B.FONT, 10, 'bold'))
        s.configure('TCheckbutton', background=B.GRAY_BG, foreground=B.GRAY_TEXT,
                    font=(B.FONT, 10))
        s.configure('TRadiobutton', background=B.GRAY_BG, foreground=B.GRAY_TEXT,
                    font=(B.FONT, 10))
        s.configure('TEntry', padding=5, fieldbackground=B.WHITE)

        s.configure('Primary.TButton', background=B.ORANGE, foreground=B.WHITE,
                    font=(B.FONT, 10, 'bold'), padding=8, borderwidth=0)
        s.map('Primary.TButton',
              background=[('active', B.ORANGE_HOVER), ('disabled', B.GRAY_BORDER)],
              foreground=[('disabled', B.GRAY_TEXT)])

        s.configure('Navy.TButton', background=B.NAVY, foreground=B.WHITE,
                    font=(B.FONT, 10), padding=8, borderwidth=0)
        s.map('Navy.TButton', background=[('active', B.NAVY_LIGHT)])

        s.configure('Browse.TButton', background=B.GRAY_BORDER, foreground=B.NAVY,
                    font=(B.FONT, 9), padding=4, borderwidth=0)
        s.map('Browse.TButton', background=[('active', '#C8C8C8')])

        s.configure('GI.Horizontal.TProgressbar',
                    troughcolor=B.GRAY_BORDER, background=B.ORANGE,
                    lightcolor=B.ORANGE, darkcolor=B.ORANGE_HOVER,
                    bordercolor=B.GRAY_BG)

        self.root.configure(background=B.GRAY_BG)
        self.root.minsize(820, 680)

    # ── UI principale ─────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main = ttk.Frame(self.root, padding=20)
        main.grid(sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(3, weight=1)

        self._build_header(main)
        self._build_params(main)
        self._build_logs(main)
        self._build_footer(main)

    def _build_header(self, parent: ttk.Frame) -> None:
        B = _Brand
        banner = tk.Frame(parent, background=B.NAVY, padx=18, pady=14)
        banner.grid(row=0, column=0, sticky="ew")
        banner.columnconfigure(1, weight=1)

        tk.Frame(banner, background=B.ORANGE, width=5).grid(
            row=0, column=0, rowspan=2, sticky="ns", padx=(0, 14))

        tk.Label(banner, text="Publipostage DOETH",
                 background=B.NAVY, foreground=B.WHITE,
                 font=(B.FONT, 17, 'bold')).grid(row=0, column=1, sticky="w")
        tk.Label(banner,
                 text="Generation automatique des attestations — Groupe Interaction",
                 background=B.NAVY, foreground="#9AB0C8",
                 font=(B.FONT, 9)).grid(row=1, column=1, sticky="w", pady=(2, 0))

        tk.Frame(parent, background=B.ORANGE, height=3).grid(
            row=1, column=0, sticky="ew", pady=(0, 14))

    def _build_params(self, parent: ttk.Frame) -> None:
        B = _Brand
        pf = ttk.LabelFrame(parent, text="Parametres", padding=12)
        pf.grid(row=2, column=0, sticky="ew")
        pf.columnconfigure(1, weight=1)

        self._row_file(pf, 0, "Fichier Excel :", self.input_file_var, self._browse_input,
                       [("Excel", "*.xlsx *.xls"), ("Tous", "*.*")])
        self._row_entry(pf, 1, "Feuille Excel :", self.sheet_var,
                        "(vide = premiere feuille)")
        self._row_dir(pf, 2, "Dossier de sortie :", self.output_dir_var, self._browse_output)

        res = ttk.LabelFrame(pf, text="Ressources", padding=10)
        res.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(10, 2))
        res.columnconfigure(1, weight=1)
        self._row_file(res, 0, "Logo :", self.logo_var, self._browse_logo,
                       [("Images", "*.png *.jpg *.jpeg")])
        self._row_file(res, 1, "Signature :", self.sig_var, self._browse_sig,
                       [("Images", "*.png *.jpg *.jpeg")])

        adv = ttk.LabelFrame(pf, text="Options avancees", padding=10)
        adv.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        adv.columnconfigure(1, weight=1)

        ttk.Checkbutton(adv, text="Ignorer le traitement Excel (utiliser un CSV existant)",
                        variable=self.skip_var, command=self._toggle_csv
                        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=3)

        ttk.Label(adv, text="Fichier CSV :").grid(row=1, column=0, sticky="w",
                                                  padx=(16, 6), pady=2)
        self.csv_entry = ttk.Entry(adv, textvariable=self.csv_path_var, state="disabled")
        self.csv_entry.grid(row=1, column=1, sticky="ew", padx=4)
        self.csv_btn = ttk.Button(adv, text="Parcourir...", style="Browse.TButton",
                                  command=self._browse_csv, state="disabled")
        self.csv_btn.grid(row=1, column=2, padx=4)

        ttk.Label(adv, text="Format de sortie :").grid(
            row=2, column=0, sticky="w", pady=(10, 3))
        fmt = ttk.Frame(adv)
        fmt.grid(row=2, column=1, columnspan=2, sticky="w", pady=(10, 3))
        for lbl, val in [("Word (.docx)", "docx"), ("PDF (.pdf)", "pdf"), ("Les deux", "both")]:
            ttk.Radiobutton(fmt, text=lbl, variable=self.fmt_var,
                            value=val).pack(side="left", padx=(0, 18))

        ttk.Checkbutton(adv, text="Mode debug (logs detailles)",
                        variable=self.debug_var).grid(
            row=3, column=0, columnspan=3, sticky="w", pady=3)

    def _build_logs(self, parent: ttk.Frame) -> None:
        lf = ttk.LabelFrame(parent, text="Journal d'execution", padding=6)
        lf.grid(row=3, column=0, sticky="nsew", pady=(10, 8))
        lf.columnconfigure(0, weight=1)
        lf.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            lf, wrap=tk.WORD, font=('Consolas', 9),
            bg="#1E1E1E", fg="#D4D4D4",
            relief="flat", borderwidth=0, state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.tag_configure("log_debug", foreground="#707070")
        self.log_text.tag_configure("log_info", foreground="#D4D4D4")
        self.log_text.tag_configure("log_warning", foreground="#E8A87C")
        self.log_text.tag_configure("log_error", foreground="#F48771")

        sb = ttk.Scrollbar(lf, orient=tk.VERTICAL, command=self.log_text.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=sb.set)

    def _build_footer(self, parent: ttk.Frame) -> None:
        B = _Brand
        footer = ttk.Frame(parent)
        footer.grid(row=4, column=0, sticky="ew")
        footer.columnconfigure(1, weight=1)

        prog = ttk.Frame(footer)
        prog.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(0, 8))
        prog.columnconfigure(1, weight=1)

        self.status_label = ttk.Label(prog, text="Pret", width=32,
                                      font=(B.FONT, 9), foreground=B.NAVY)
        self.status_label.grid(row=0, column=0, sticky="w")

        self.progress_var = tk.DoubleVar()
        ttk.Progressbar(prog, style='GI.Horizontal.TProgressbar',
                        variable=self.progress_var, maximum=100
                        ).grid(row=0, column=1, sticky="ew", padx=8)

        self.pct_label = ttk.Label(prog, text="0%", width=5,
                                   font=(B.FONT, 9, 'bold'), foreground=B.ORANGE)
        self.pct_label.grid(row=0, column=2, sticky="e")

        btn = ttk.Frame(footer)
        btn.grid(row=1, column=0, columnspan=4, sticky="ew")
        btn.columnconfigure((0, 1, 2), weight=1)

        self.start_btn = ttk.Button(btn, text="Demarrer le traitement",
                                    style="Primary.TButton",
                                    command=self._start_processing)
        self.start_btn.grid(row=0, column=0, padx=(0, 6), sticky="ew")

        ttk.Button(btn, text="Ouvrir le dossier de sortie",
                   style="Navy.TButton",
                   command=self._open_output).grid(row=0, column=1, padx=6, sticky="ew")

        ttk.Button(btn, text="Quitter", style="Browse.TButton",
                   command=self._on_close).grid(row=0, column=2, padx=(6, 0), sticky="ew")

    # ── Helpers construction ───────────────────────────────────────────────────
    def _row_file(self, parent, row, label, var, cmd, filetypes) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w",
                                           padx=(0, 6), pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew",
                                                 padx=4, pady=4)
        ttk.Button(parent, text="Parcourir...", style="Browse.TButton",
                   command=cmd).grid(row=row, column=2, pady=4)

    def _row_dir(self, parent, row, label, var, cmd) -> None:
        self._row_file(parent, row, label, var, cmd, [])

    def _row_entry(self, parent, row, label, var, hint="") -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w",
                                           padx=(0, 6), pady=4)
        ttk.Entry(parent, textvariable=var, width=26).grid(row=row, column=1,
                                                           sticky="w", padx=4, pady=4)
        if hint:
            ttk.Label(parent, text=hint,
                      font=(_Brand.FONT, 8), foreground="#909090"
                      ).grid(row=row, column=2, sticky="w", padx=4)

    # ── Parcourir ──────────────────────────────────────────────────────────────
    def _browse_input(self) -> None:
        p = filedialog.askopenfilename(title="Selectionner le fichier Excel",
                                       filetypes=[("Excel", "*.xlsx *.xls"), ("Tous", "*.*")])
        if p:
            self.input_file_var.set(p)

    def _browse_output(self) -> None:
        p = filedialog.askdirectory(title="Dossier de sortie")
        if p:
            self.output_dir_var.set(p)

    def _browse_csv(self) -> None:
        p = filedialog.askopenfilename(title="Selectionner le CSV",
                                       filetypes=[("CSV", "*.csv"), ("Tous", "*.*")])
        if p:
            self.csv_path_var.set(p)

    def _browse_logo(self) -> None:
        p = filedialog.askopenfilename(title="Selectionner le logo",
                                       filetypes=[("Images", "*.png *.jpg *.jpeg")])
        if p:
            self.logo_var.set(p)

    def _browse_sig(self) -> None:
        p = filedialog.askopenfilename(title="Selectionner la signature",
                                       filetypes=[("Images", "*.png *.jpg *.jpeg")])
        if p:
            self.sig_var.set(p)

    def _toggle_csv(self) -> None:
        s = "normal" if self.skip_var.get() else "disabled"
        self.csv_entry.configure(state=s)
        self.csv_btn.configure(state=s)

    # ── Logging ───────────────────────────────────────────────────────────────
    def _setup_logging(self) -> None:
        handler = _LogHandler(self.log_text)
        handler.setFormatter(logging.Formatter(
            '%(asctime)s  %(levelname)-8s  %(message)s', '%H:%M:%S'))
        root_log = logging.getLogger()
        root_log.setLevel(logging.DEBUG)
        for h in root_log.handlers[:]:
            root_log.removeHandler(h)
        root_log.addHandler(handler)
        self.logger = logging.getLogger("gui")
        self.logger.info("Interface demarre — Publipostage DOETH")
        for name, var in [("Logo", self.logo_var), ("Signature", self.sig_var)]:
            path = var.get()
            if Path(path).exists():
                self.logger.info(f"{name} trouve : {path}")
            else:
                self.logger.warning(f"{name} non trouve : {path}")

    # ── Progression (thread-safe via after) ───────────────────────────────────
    def _set_progress(self, value: float, status: str = "") -> None:
        def _upd() -> None:
            self.progress_var.set(value)
            self.pct_label.configure(text=f"{int(value)}%")
            if status:
                self.status_label.configure(text=status)
        self.root.after(0, _upd)

    # ── Traitement ────────────────────────────────────────────────────────────
    def _start_processing(self) -> None:
        if self.processing:
            messagebox.showinfo("En cours", "Un traitement est deja en cours.")
            return
        if not self.skip_var.get() and not Path(self.input_file_var.get()).exists():
            messagebox.showerror("Erreur", "Fichier Excel introuvable.")
            return
        if self.skip_var.get() and not Path(self.csv_path_var.get()).exists():
            messagebox.showerror("Erreur", "Fichier CSV introuvable.")
            return
        for label, var in [("logo", self.logo_var), ("signature", self.sig_var)]:
            if not Path(var.get()).exists():
                if not messagebox.askyesno("Ressource manquante",
                                           f"Le {label} est introuvable.\nContinuer ?"):
                    return

        fmt_map = {"docx": OutputFormat.DOCX, "pdf": OutputFormat.PDF,
                   "both": OutputFormat.BOTH}
        args = {
            "input": self.input_file_var.get(),
            "sheet": self.sheet_var.get(),
            "output_dir": self.output_dir_var.get(),
            "skip_processing": self.skip_var.get(),
            "csv_path": self.csv_path_var.get() if self.skip_var.get() else None,
            "logo_path": self.logo_var.get(),
            "signature_path": self.sig_var.get(),
            "debug": self.debug_var.get(),
            "output_format": fmt_map.get(self.fmt_var.get(), OutputFormat.DOCX),
        }

        self.processing = True
        self.start_btn.configure(state="disabled")
        self._set_progress(0, "Demarrage...")

        threading.Thread(target=self._worker, args=(args,), daemon=True).start()
        self.root.after(200, self._poll)

    def _worker(self, args: dict) -> None:
        """
        Exécute la pipeline dans un thread secondaire.

        Utilise run_pipeline() de main.py comme point d'entrée unique,
        éliminant la duplication d'orchestration entre CLI et GUI.
        """
        try:
            from main import setup_environment, run_pipeline

            self._set_progress(5, "Initialisation...")

            # Construit un Namespace compatible avec setup_environment
            # sans passer par _NS (code smell) ni dupliquer la logique de validation
            ns = argparse.Namespace(
                config=None,
                input=args['input'],
                sheet=args['sheet'],
                output_dir=args['output_dir'],
                skip_processing=args['skip_processing'],
                csv_path=args['csv_path'],
                debug=args['debug'],
                dry_run=False,
            )
            params, _ = setup_environment(ns)

            # Surcharge logo/sig depuis les champs GUI (priorité sur la config)
            if args['logo_path'] and Path(args['logo_path']).exists():
                params['logo_path'] = args['logo_path']
            if args['signature_path'] and Path(args['signature_path']).exists():
                params['signature_path'] = args['signature_path']
            params['output_format'] = args['output_format']

            # Le GUI utilise son propre logger (propagation root → _LogHandler → widget)
            # et non le session_logger de setup_environment (propagate=False)
            docs, stats = run_pipeline(
                params,
                self.logger,
                progress_callback=self._set_progress,
            )

            self.logger.info("─── BILAN ───────────────────────────────")
            self.logger.info(f"  Attestations  : {stats.total_docs}")
            self.logger.info(f"  SIRET traites : {stats.unique_sirets}")
            if stats.total_etp:
                self.logger.info(f"  Total ETP     : {stats.total_etp:.2f}")
            self.logger.info(f"  Dossier       : {args['output_dir']}")
            self.logger.info("─── TERMINE AVEC SUCCES ─────────────────")

        except Exception as e:
            import traceback
            self.logger.error(f"Erreur : {e}")
            self.logger.error(traceback.format_exc())
            self._set_progress(100, "Erreur — voir le journal")
        finally:
            self.processing = False

    def _poll(self) -> None:
        if self.processing:
            self.root.after(200, self._poll)
            return
        self.start_btn.configure(state="normal")
        txt = self.status_label.cget("text")
        if "Termine" in txt and "Erreur" not in txt:
            if messagebox.askyesno("Succes", "Traitement termine !\nOuvrir le dossier ?"):
                self._open_output()

    def _open_output(self) -> None:
        d = self.output_dir_var.get()
        if not Path(d).exists():
            messagebox.showwarning("Introuvable", f"Dossier introuvable :\n{d}")
            return
        try:
            if sys.platform == 'win32':
                import os
                os.startfile(d)
            elif sys.platform == 'darwin':
                result = subprocess.run(['open', d], capture_output=True)
                if result.returncode != 0:
                    self.logger.warning(f"Impossible d'ouvrir le dossier : {d}")
            else:
                result = subprocess.run(['xdg-open', d], capture_output=True)
                if result.returncode != 0:
                    self.logger.warning(f"Impossible d'ouvrir le dossier : {d}")
        except Exception as e:
            self.logger.error(f"Impossible d'ouvrir : {e}")

    def _on_close(self) -> None:
        if self.processing:
            if not messagebox.askyesno("Traitement en cours",
                                       "Un traitement est en cours.\nQuitter ?"):
                return
        self.root.destroy()


# ── Point d'entrée ─────────────────────────────────────────────────────────────
def main() -> None:
    root = tk.Tk()
    PublipostageGUI(root)
    try:
        ico = Path(__file__).parent / 'resources' / 'icon.ico'
        if ico.exists():
            root.iconbitmap(str(ico))
    except Exception:
        pass
    w, h = 900, 820
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")
    root.mainloop()


if __name__ == '__main__':
    main()
