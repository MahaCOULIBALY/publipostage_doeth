#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interface graphique modernisée pour le Publipostage DOETH.
L'IHM a été optimisée pour une expérience utilisateur optimale : responsive,
palette de couleurs fonctionnelles et indicateurs de progression clairs.
"""

import os
import sys
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import queue
import logging

# Importer les modules du projet
from src.utils.config import get, Config
from src.utils.logger import setup_logger, get_logger
from src.data_processor import nettoyer_fichier_excel
from src.document_generator import generer_attestations_doeth


class RedirectText:
    """Redirige les logs vers un widget Text avec file d'attente."""

    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.queue = queue.Queue()
        self.update_timer = None

    def write(self, string):
        self.queue.put(string)
        if self.update_timer is None:
            self.update_timer = self.text_widget.after(
                100, self.update_text_widget)

    def update_text_widget(self):
        self.update_timer = None
        try:
            while True:
                string = self.queue.get_nowait()
                self.text_widget.configure(state="normal")
                self.text_widget.insert("end", string)
                self.text_widget.see("end")
                self.text_widget.configure(state="disabled")
                self.queue.task_done()
        except queue.Empty:
            pass

    def flush(self):
        pass


class LoggingHandler(logging.Handler):
    """Handler personnalisé pour rediriger les logs avec coloration."""

    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.level_colors = {
            logging.DEBUG: "gray",
            logging.INFO: "black",
            logging.WARNING: "orange",
            logging.ERROR: "red",
            logging.CRITICAL: "red"
        }

    def emit(self, record):
        msg = self.format(record)
        level_color = self.level_colors.get(record.levelno, "black")

        def update_log():
            self.text_widget.configure(state="normal")
            self.text_widget.insert("end", msg + "\n", level_color)
            self.text_widget.see("end")
            self.text_widget.configure(state="disabled")

        self.text_widget.after(0, update_log)


class PublipostageGUI:
    """Interface graphique modernisée pour le publipostage DOETH."""

    def __init__(self, root):
        self.root = root
        root.title("Publipostage DOETH")

        # Palette de couleurs et paramètres de design
        self.primary_color = "#1976D2"  # Bleu pour actions d'ouverture
        self.start_color = "#388E3C"  # Vert pour démarrer
        self.exit_color = "#D32F2F"  # Rouge pour quitter
        self.bg_color = "#F5F5F5"  # Fond gris très clair
        self.text_color = "#212121"  # Texte gris foncé

        # Configuration du thème et styles
        style = ttk.Style()
        if 'clam' in style.theme_names():
            style.theme_use('clam')

        style.configure('TFrame', background=self.bg_color)
        style.configure('TLabel', background=self.bg_color,
                        foreground=self.text_color, font=('Segoe UI', 10))
        style.configure('TLabelframe', background=self.bg_color,
                        foreground=self.text_color)
        style.configure('TLabelframe.Label', background=self.bg_color, foreground=self.primary_color,
                        font=('Segoe UI', 11, 'bold'))
        style.configure('TEntry', padding=5)
        style.configure('TCheckbutton', background=self.bg_color,
                        foreground=self.text_color, font=('Segoe UI', 10))

        # Boutons avec styles spécifiques
        style.configure('Start.TButton', background=self.start_color, foreground="white", font=('Segoe UI', 10, 'bold'),
                        padding=6, borderwidth=0)
        style.map('Start.TButton', background=[('active', "#66BB6A")])
        style.configure('Open.TButton', background=self.primary_color, foreground="white",
                        font=('Segoe UI', 10, 'bold'), padding=6, borderwidth=0)
        style.map('Open.TButton', background=[('active', "#64B5F6")])
        style.configure('Exit.TButton', background=self.exit_color, foreground="white", font=('Segoe UI', 10, 'bold'),
                        padding=6, borderwidth=0)
        style.map('Exit.TButton', background=[('active', "#E57373")])

        # Barre de progression personnalisée (vert)
        style.configure("Green.Horizontal.TProgressbar", troughcolor=self.bg_color, bordercolor=self.bg_color,
                        background="#388E3C", lightcolor="#66BB6A", darkcolor="#2E7D32")

        # Configuration de la fenêtre
        root.geometry("900x700")
        root.minsize(700, 500)
        root.configure(background=self.bg_color)

        # Variables pour les paramètres
        self.input_file_var = tk.StringVar(value=os.path.join(
            get('paths.input_dir', './data/input'),
            get('defaults.input_filename', 'donnees.xlsx')
        ))
        self.sheet_name_var = tk.StringVar(
            value=get('defaults.excel_sheet', 'Feuil1'))
        self.output_dir_var = tk.StringVar(
            value=get('paths.output_dir', './data/output'))
        self.csv_path_var = tk.StringVar()
        self.logo_path_var = tk.StringVar(value=get('resources.logo_path', ''))
        self.signature_path_var = tk.StringVar(
            value=get('resources.signature_path', ''))
        self.skip_processing_var = tk.BooleanVar(value=False)
        self.debug_var = tk.BooleanVar(value=False)
        self.output_format_var = tk.StringVar(value="docx")

        # Etat de traitement
        self.processing = False
        self.process_thread = None

        self.create_widgets()
        self.setup_logging()
        root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        """Création et organisation des widgets avec un design responsive et coloré."""
        main_frame = ttk.Frame(self.root, padding="20", style="TFrame")
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        # En-tête
        header_frame = ttk.Frame(main_frame, style="TFrame")
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        title_label = ttk.Label(
            header_frame, text="Publipostage DOETH", font=("Segoe UI", 20, "bold"))
        title_label.grid(row=0, column=0, sticky="w")
        description_label = ttk.Label(header_frame,
                                      text="Génération d'attestations pour travailleurs en situation de handicap",
                                      font=("Segoe UI", 12))
        description_label.grid(row=1, column=0, sticky="w", pady=(5, 0))

        # Séparateur
        separator = ttk.Separator(main_frame, orient="horizontal")
        separator.grid(row=1, column=0, sticky="ew", pady=10)

        # Section Paramètres
        params_frame = ttk.LabelFrame(
            main_frame, text="Paramètres", padding="15", style="TLabelframe")
        params_frame.grid(row=2, column=0, sticky="ew", pady=10)
        params_frame.columnconfigure(1, weight=1)

        # Paramètres principaux
        ttk.Label(params_frame, text="Fichier Excel:").grid(
            row=0, column=0, sticky="w", padx=5, pady=5)
        input_entry = ttk.Entry(params_frame, textvariable=self.input_file_var)
        input_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(params_frame, text="Parcourir...", command=self.browse_input_file).grid(row=0, column=2, padx=5,
                                                                                           pady=5)

        ttk.Label(params_frame, text="Feuille Excel:").grid(
            row=1, column=0, sticky="w", padx=5, pady=5)
        sheet_entry = ttk.Entry(
            params_frame, textvariable=self.sheet_name_var, width=20)
        sheet_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(params_frame, text="(laisser vide pour la première)", font=("Segoe UI", 9), foreground="gray") \
            .grid(row=1, column=2, sticky="w", padx=5, pady=5)

        ttk.Label(params_frame, text="Dossier de sortie:").grid(
            row=2, column=0, sticky="w", padx=5, pady=5)
        output_entry = ttk.Entry(
            params_frame, textvariable=self.output_dir_var)
        output_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(params_frame, text="Parcourir...", command=self.browse_output_dir).grid(row=2, column=2, padx=5,
                                                                                           pady=5)

        # Paramètres Logo et Signature
        resources_frame = ttk.LabelFrame(params_frame, text="Ressources des attestations", padding="15",
                                         style="TLabelframe")
        resources_frame.grid(row=3, column=0, columnspan=3,
                             sticky="ew", padx=5, pady=(15, 5))
        resources_frame.columnconfigure(1, weight=1)

        ttk.Label(resources_frame, text="Logo:").grid(
            row=0, column=0, sticky="w", padx=5, pady=5)
        logo_entry = ttk.Entry(
            resources_frame, textvariable=self.logo_path_var)
        logo_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(resources_frame, text="Parcourir...", command=self.browse_logo_file).grid(row=0, column=2, padx=5,
                                                                                             pady=5)

        ttk.Label(resources_frame, text="Signature:").grid(
            row=1, column=0, sticky="w", padx=5, pady=5)
        signature_entry = ttk.Entry(
            resources_frame, textvariable=self.signature_path_var)
        signature_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(resources_frame, text="Parcourir...", command=self.browse_signature_file).grid(row=1, column=2,
                                                                                                  padx=5, pady=5)

        # Options avancées
        advanced_frame = ttk.LabelFrame(
            params_frame, text="Options avancées", padding="15", style="TLabelframe")
        advanced_frame.grid(row=4, column=0, columnspan=3,
                            sticky="ew", padx=5, pady=15)
        advanced_frame.columnconfigure(1, weight=1)
        ttk.Checkbutton(advanced_frame, text="Ignorer le traitement Excel",
                        variable=self.skip_processing_var, command=self.toggle_csv_path) \
            .grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
        ttk.Label(advanced_frame, text="Fichier CSV:").grid(
            row=1, column=0, sticky="w", padx=5, pady=5)
        self.csv_entry = ttk.Entry(
            advanced_frame, textvariable=self.csv_path_var, state="disabled")
        self.csv_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        self.csv_button = ttk.Button(advanced_frame, text="Parcourir...", command=self.browse_csv_file,
                                     state="disabled")
        self.csv_button.grid(row=1, column=2, padx=5, pady=5)
        ttk.Checkbutton(advanced_frame, text="Mode debug (logs détaillés)", variable=self.debug_var) \
            .grid(row=2, column=0, columnspan=2, sticky="w", pady=5)
        # Format de sortie
        ttk.Label(advanced_frame, text="Format de sortie :").grid(
            row=3, column=0, sticky="w", padx=5, pady=5)
        format_frame = ttk.Frame(advanced_frame)
        format_frame.grid(row=3, column=1, columnspan=2,
                          sticky="w", padx=5, pady=5)
        for label, value in [("Word (.docx)", "docx"), ("PDF (.pdf)", "pdf"), ("Les deux", "both")]:
            ttk.Radiobutton(format_frame, text=label, variable=self.output_format_var, value=value) \
                .pack(side="left", padx=10)

        # Zone de détails des traitements (anciennement "Logs")
        log_frame = ttk.LabelFrame(
            main_frame, text="Détails des traitements", padding="15", style="TLabelframe")
        log_frame.grid(row=4, column=0, sticky="nsew", pady=10)
        main_frame.rowconfigure(4, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, wrap=tk.WORD,
                                font=('Consolas', 9), bg="white")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.tag_configure("gray", foreground="#707070")
        self.log_text.tag_configure("black", foreground="#000000")
        self.log_text.tag_configure("orange", foreground="#FF8C00")
        self.log_text.tag_configure("red", foreground="#FF0000")
        self.log_text.configure(state="disabled")
        log_scrollbar = ttk.Scrollbar(
            log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scrollbar.set)

        # Barre d'état et de progression
        status_frame = ttk.Frame(main_frame, style="TFrame")
        status_frame.grid(row=5, column=0, sticky="ew", pady=(10, 15))
        self.status_label = ttk.Label(
            status_frame, text="Prêt", font=("Segoe UI", 10))
        self.status_label.grid(row=0, column=0, sticky="w", padx=5)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(status_frame, style="Green.Horizontal.TProgressbar",
                                            variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=1, sticky="ew", padx=5)
        status_frame.columnconfigure(1, weight=1)
        # Label pour afficher le pourcentage sur la jauge
        self.progress_percentage = ttk.Label(
            status_frame, text="0%", font=("Segoe UI", 10), background=self.bg_color)
        self.progress_percentage.grid(row=0, column=2, padx=5)

        # Boutons d'action (responsive)
        buttons_frame = ttk.Frame(main_frame, style="TFrame")
        buttons_frame.grid(row=6, column=0, sticky="ew", pady=10)
        buttons_frame.columnconfigure((0, 1, 2), weight=1)
        ttk.Button(buttons_frame, text="Démarrer le traitement", command=self.start_processing, style="Start.TButton") \
            .grid(row=0, column=0, padx=10, sticky="ew")
        ttk.Button(buttons_frame, text="Ouvrir dossier de sortie", command=self.open_output_folder,
                   style="Open.TButton") \
            .grid(row=0, column=1, padx=10, sticky="ew")
        ttk.Button(buttons_frame, text="Quitter", command=self.on_closing, style="Exit.TButton") \
            .grid(row=0, column=2, padx=10, sticky="ew")

    def setup_logging(self):
        """Configure le logging pour rediriger vers le widget Text."""
        text_handler = LoggingHandler(self.log_text)
        text_handler.setLevel(logging.DEBUG)
        formatter = logging.Formatter(
            '%(asctime)s | %(levelname)-8s | %(message)s', '%H:%M:%S')
        text_handler.setFormatter(formatter)
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.DEBUG)
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)
        root_logger.addHandler(text_handler)
        self.logger = logging.getLogger("publipostage_gui")
        self.logger.info("Interface graphique démarrée")

        # Vérification des ressources
        logo_path = self.logo_path_var.get()
        signature_path = self.signature_path_var.get()

        if not os.path.exists(logo_path):
            self.logger.warning(f"Logo non trouvé: {logo_path}")
        else:
            self.logger.info(f"Logo trouvé: {logo_path}")

        if not os.path.exists(signature_path):
            self.logger.warning(f"Signature non trouvée: {signature_path}")
        else:
            self.logger.info(f"Signature trouvée: {signature_path}")

    def toggle_csv_path(self):
        """Active ou désactive les champs relatifs au fichier CSV."""
        if self.skip_processing_var.get():
            self.csv_entry.configure(state="normal")
            self.csv_button.configure(state="normal")
        else:
            self.csv_entry.configure(state="disabled")
            self.csv_button.configure(state="disabled")

    def browse_input_file(self):
        filepath = filedialog.askopenfilename(
            title="Sélectionner le fichier Excel",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls"),
                       ("Tous les fichiers", "*.*")]
        )
        if filepath:
            self.input_file_var.set(filepath)

    def browse_output_dir(self):
        dirpath = filedialog.askdirectory(
            title="Sélectionner le dossier de sortie")
        if dirpath:
            self.output_dir_var.set(dirpath)

    def browse_csv_file(self):
        filepath = filedialog.askopenfilename(
            title="Sélectionner le fichier CSV",
            filetypes=[("Fichiers CSV", "*.csv"), ("Tous les fichiers", "*.*")]
        )
        if filepath:
            self.csv_path_var.set(filepath)

    def browse_logo_file(self):
        filepath = filedialog.askopenfilename(
            title="Sélectionner le logo",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.bmp"),
                       ("Tous les fichiers", "*.*")]
        )
        if filepath:
            self.logo_path_var.set(filepath)
            if os.path.exists(filepath):
                self.logger.info(f"Logo sélectionné: {filepath}")

    def browse_signature_file(self):
        filepath = filedialog.askopenfilename(
            title="Sélectionner la signature",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.bmp"),
                       ("Tous les fichiers", "*.*")]
        )
        if filepath:
            self.signature_path_var.set(filepath)
            if os.path.exists(filepath):
                self.logger.info(f"Signature sélectionnée: {filepath}")

    def open_output_folder(self):
        output_dir = self.output_dir_var.get()
        if not os.path.exists(output_dir):
            messagebox.showwarning(
                "Attention", f"Le dossier de sortie n'existe pas : {output_dir}")
            return
        try:
            if sys.platform == 'win32':
                os.startfile(output_dir)
            elif sys.platform == 'darwin':
                subprocess.run(['open', output_dir])
            else:
                subprocess.run(['xdg-open', output_dir])
            self.logger.info(f"Dossier ouvert : {output_dir}")
        except Exception as e:
            self.logger.error(
                f"Erreur lors de l'ouverture du dossier : {str(e)}")
            messagebox.showerror(
                "Erreur", f"Impossible d'ouvrir le dossier : {str(e)}")

    def start_processing(self):
        """Lance le traitement dans un thread séparé après vérification des paramètres."""
        if self.processing:
            messagebox.showinfo(
                "Information", "Un traitement est déjà en cours.")
            return
        if not self.skip_processing_var.get() and not os.path.exists(self.input_file_var.get()):
            messagebox.showerror(
                "Erreur", "Le fichier Excel d'entrée n'existe pas.")
            return
        if self.skip_processing_var.get() and not os.path.exists(self.csv_path_var.get()):
            messagebox.showerror(
                "Erreur", "Le fichier CSV spécifié n'existe pas.")
            return

        # Vérifier les ressources logo et signature
        if not os.path.exists(self.logo_path_var.get()):
            if not messagebox.askyesno("Attention",
                                       "Le logo n'existe pas ou n'a pas été spécifié. Voulez-vous continuer quand même ?"):
                return

        if not os.path.exists(self.signature_path_var.get()):
            if not messagebox.askyesno("Attention",
                                       "La signature n'existe pas ou n'a pas été spécifiée. Voulez-vous continuer quand même ?"):
                return

        os.makedirs(self.output_dir_var.get(), exist_ok=True)
        self.processing = True
        self.disable_buttons(True)
        self.progress_var.set(0)
        self.update_progress_percentage(0)
        args = {
            "input": self.input_file_var.get(),
            "sheet": self.sheet_name_var.get(),
            "output_dir": self.output_dir_var.get(),
            "skip_processing": self.skip_processing_var.get(),
            "csv_path": self.csv_path_var.get() if self.skip_processing_var.get() else None,
            "logo_path": self.logo_path_var.get(),
            "signature_path": self.signature_path_var.get(),
            "debug": self.debug_var.get(),
            "output_format": self.output_format_var.get()
        }
        self.process_thread = threading.Thread(
            target=self.run_processing_thread, args=(args,), daemon=True)
        self.process_thread.start()
        self.root.after(100, self.check_process_status)

    def disable_buttons(self, processing):
        """Active ou désactive les boutons d'action en fonction de l'état."""
        state = "disabled" if processing else "normal"
        # Pour chaque bouton dans le conteneur des boutons
        for child in self.root.winfo_children():
            try:
                child.configure(state=state)
            except:
                pass

    def run_processing_thread(self, args):
        self.logger.info("Démarrage du traitement avec les paramètres :")
        self.logger.info(f"  Fichier Excel : {args['input']}")
        self.logger.info(f"  Feuille : {args['sheet']}")
        self.logger.info(f"  Dossier de sortie : {args['output_dir']}")
        self.logger.info(f"  Logo : {args['logo_path']}")
        self.logger.info(f"  Signature : {args['signature_path']}")
        try:
            from main import setup_environment, generate_statistics
            import pandas as pd
            import time

            self.update_progress(5, "Configuration de l'environnement...")

            class Args:
                def __init__(self, **kwargs):
                    for key, value in kwargs.items():
                        setattr(self, key, value)

            args_obj = Args(**args)
            params = setup_environment(args_obj)

            # Utiliser les chemins de logo et signature spécifiés dans l'interface
            if args['logo_path'] and os.path.exists(args['logo_path']):
                params['logo_path'] = args['logo_path']

            if args['signature_path'] and os.path.exists(args['signature_path']):
                params['signature_path'] = args['signature_path']

            app_logger = get_logger()

            self.update_progress(10, "Traitement des données Excel...")
            csv_path = ""
            if args['skip_processing']:
                csv_path = args['csv_path']
                self.logger.info(f"Utilisation du CSV existant : {csv_path}")
            else:
                start_time = time.time()
                self.logger.info(
                    f"Traitement du fichier Excel : {args['input']}")
                df_processed = nettoyer_fichier_excel(
                    input_file=args['input'],
                    output_file=params['csv_path'],
                    sheet_name=args['sheet'],
                    logger=app_logger
                )
                csv_path = params['csv_path']
                elapsed_time = time.time() - start_time
                self.logger.info(
                    f"Traitement terminé en {elapsed_time:.2f} sec")
                self.logger.info(
                    f"Lignes traitées : {len(df_processed)} ; Colonnes : {df_processed.columns.size}")
                self.logger.info(
                    f"SIRET uniques : {df_processed['SIRET'].nunique()}")
            self.update_progress(40, "CSV créé avec succès")
            separator = get('defaults.csv_separator', ';')
            df_processed = pd.read_csv(csv_path, sep=separator, quoting=1)
            self.update_progress(50, "Génération des attestations...")
            start_time = time.time()
            from document_generator import OutputFormat
            fmt_map = {"docx": OutputFormat.DOCX,
                       "pdf": OutputFormat.PDF, "both": OutputFormat.BOTH}
            output_fmt = fmt_map.get(
                args.get("output_format", "docx"), OutputFormat.DOCX)
            generated_docs = generer_attestations_doeth(
                csv_path=csv_path,
                output_folder=args['output_dir'],
                logger=app_logger,
                signature_path=params['signature_path'],
                logo_path=params['logo_path'],
                output_format=output_fmt,
            )
            elapsed_time = time.time() - start_time
            self.logger.info(
                f"Attestations générées en {elapsed_time:.2f} sec : {len(generated_docs)} documents")
            self.update_progress(85, "Attestations générées")
            self.update_progress(90, "Calcul des statistiques...")
            try:
                stats = generate_statistics(df_processed, generated_docs)
            except:
                stats = {
                    "total_rows": len(df_processed),
                    "unique_sirets": df_processed['SIRET'].nunique(),
                    "unique_clients": df_processed[
                        'NOM_CLIENT'].nunique() if 'NOM_CLIENT' in df_processed.columns else 0,
                    "total_docs": len(generated_docs)
                }
                if 'ETP_ANNUEL' in df_processed.columns:
                    stats["total_etp"] = df_processed['ETP_ANNUEL'].sum()
            self.update_progress(95, "Finalisation...")
            self.logger.info("=== BILAN DU TRAITEMENT ===")
            self.logger.info(f"Total attestations : {len(generated_docs)}")
            self.logger.info(
                f"SIRET traités : {stats.get('unique_sirets', 'N/A')}")
            if 'unique_clients' in stats:
                self.logger.info(
                    f"Clients uniques : {stats['unique_clients']}")
            if 'total_etp' in stats:
                self.logger.info(f"Total ETP : {stats['total_etp']:.2f}")
            self.logger.info(f"Dossier de sortie : {args['output_dir']}")
            self.logger.info("=== TRAITEMENT TERMINÉ AVEC SUCCÈS ===")
            self.update_progress(100, "Traitement terminé")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement : {str(e)}")
            import traceback
            self.logger.error(traceback.format_exc())
            self.update_progress(100, "Erreur lors du traitement")
        finally:
            self.processing = False

    def update_progress(self, value, status_text=None):
        def update():
            self.progress_var.set(value)
            self.update_progress_percentage(value)
            if status_text:
                self.status_label.configure(text=status_text)

        self.root.after(0, update)

    def update_progress_percentage(self, value):
        """Met à jour le label du pourcentage de la barre de progression."""
        self.progress_percentage.configure(text=f"{int(value)}%")

    def check_process_status(self):
        if not self.processing:
            self.disable_buttons(False)
            if self.progress_var.get() == 100:
                if "Erreur" not in self.status_label.cget("text"):
                    messagebox.showinfo(
                        "Information", "Traitement terminé avec succès!")
                    if messagebox.askyesno("Information", "Ouvrir le dossier des attestations générées ?"):
                        self.open_output_folder()
            return
        self.root.after(100, self.check_process_status)

    def on_closing(self):
        if self.processing:
            if messagebox.askyesno("Confirmation", "Un traitement est en cours. Quitter malgré tout ?"):
                self.root.destroy()
        else:
            self.root.destroy()


def main():
    root = tk.Tk()
    app = PublipostageGUI(root)
    try:
        icon_path = os.path.join(os.path.dirname(
            os.path.abspath(__file__)), 'resources', 'icon.ico')
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except:
        pass

    window_width = 850
    window_height = 950
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int((screen_width - window_width) / 2)
    center_y = int((screen_height - window_height) / 2)
    root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
    root.mainloop()


if __name__ == '__main__':
    main()
