import os
import sys
import json
from datetime import datetime

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from process_surf import (
    extract_info, tab_cd_type, build_affectation_mapping,
    TCD2Tab, Tab_output, export_tables_to_excel, export_tables_to_html,
    html_from_excel, update_glossary,
)

SVFILL_VERSION = 1
APP_NAME    = "Rfill"
APP_TITLE   = "Rfill — Analyse des Surfaces"
APP_VERSION = "1.3.0"
APP_BUILD   = "2026.06.01"
APP_DATE    = "1 juin 2026"
APP_AUTHOR  = "J. FAGUET"
APP_COMPANY = "RTaxes"
APP_DESC    = (
    "Importation et catégorisation de surfaces AutoCAD/GeoGex,\n"
    "génération de tableaux Excel normés (SUB, SU, SHO, SDP…)."
)
APP_COPY    = "© 2026 RTaxes — Tous droits réservés"
APP_LEGAL   = (
    "Ce logiciel est la propriété exclusive de RTaxes. "
    "Toute reproduction, distribution ou utilisation non autorisée, "
    "en tout ou en partie, est strictement interdite sans l'accord écrit préalable de RTaxes. "
    "Ce logiciel est fourni « tel quel », sans garantie d'aucune sorte."
)


class SurfaceApp(tk.Tk):
    def __init__(self):
        """Initialise la fenêtre principale, les styles ttk et les 5 onglets."""
        super().__init__()

        self.title(APP_TITLE)
        self.geometry("1200x780")
        self.configure(bg="#f0f0f0")

        _base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        _ico  = os.path.join(_base, "Rfill.ico")
        if os.path.isfile(_ico):
            self.iconbitmap(_ico)

        self._setup_styles()

        # ── Barre de session (fichier .svfill courant) ──────────────────────
        self._session_path = None
        self._build_session_bar()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)

        self.tab_import    = ttk.Frame(self.notebook)
        self.tab_glossaire = ttk.Frame(self.notebook)
        self.tab_generer   = ttk.Frame(self.notebook)
        self.tab_aide      = ttk.Frame(self.notebook)
        self.tab_apropos   = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_import,    text="  1. Importer  ")
        self.notebook.add(self.tab_glossaire, text="  2. Glossaire  ")
        self.notebook.add(self.tab_generer,   text="  3. Générer  ")
        self.notebook.add(self.tab_aide,      text="  Aide  ")
        self.notebook.add(self.tab_apropos,   text="  À propos  ")
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        self.all_df           = None
        self.df_types         = None
        self.mappings         = {}
        self.empty_categories = {}
        self.super_cats       = {}
        self.saved_glossaries = set()
        self.columns          = {}
        self.cat_colors       = {}   # {type_su: {cat_name: color}}
        self._loaded_files    = []   # [{"path": ..., "source": "autocad"|"geogex"}]

        self.drag_item       = None
        self.drag_label      = None
        self.drag_origin_cat = None

        self._cat_drag       = None   # {'cat', 'sc', 'start_x', 'start_y', 'active'}
        self._cat_drag_label = None

        self.build_import_tab()
        self.build_glossaire_tab()
        self.build_generer_tab()
        self.build_aide_tab()
        self.build_apropos_tab()

    # ------------------------------------------------------------------ styles

    def _setup_styles(self):
        """Configure le thème clam et les styles ttk (polices, couleurs)."""
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TNotebook",           background="#f0f0f0")
        s.configure("TNotebook.Tab",       font=("Arial", 10), padding=[10, 4])
        s.configure("TFrame",              background="#f0f0f0")
        s.configure("TLabel",              background="#f0f0f0", font=("Arial", 10))
        s.configure("Title.TLabel",        font=("Arial", 13, "bold"), foreground="#2c3e50")
        s.configure("Sub.TLabel",          font=("Arial", 9),          foreground="#666666")
        s.configure("TButton",             font=("Arial", 10),          padding=[8, 4])
        s.configure("Primary.TButton",     font=("Arial", 10, "bold"),  padding=[10, 5])

    # ---------------------------------------------------------------- session bar

    def _build_session_bar(self):
        """Barre persistante en haut : nom du fichier .svfill courant + boutons Nouveau/Ouvrir/Enregistrer."""
        bar = tk.Frame(self, bg="#2c3e50", pady=4)
        bar.pack(fill="x", side="top")

        self._session_label = tk.Label(
            bar, text="Aucune session ouverte",
            font=("Arial", 9), bg="#2c3e50", fg="#aabbcc", anchor="w",
        )
        self._session_label.pack(side="left", padx=10)

        for text, cmd in [
            ("Nouveau",       self.session_new),
            ("Ouvrir…",       self.session_open),
            ("Enregistrer",   self.session_save),
            ("Enregistrer sous…", self.session_save_as),
        ]:
            tk.Button(
                bar, text=text, command=cmd,
                font=("Arial", 8), bg="#3d5166", fg="white",
                relief="flat", padx=8, pady=2, cursor="hand2",
                activebackground="#4a6380", activeforeground="white",
            ).pack(side="right", padx=2)

    def _update_session_label(self):
        if self._session_path:
            name = os.path.basename(self._session_path)
            self._session_label.config(text=f"Session : {name}", fg="#98c379")
        else:
            self._session_label.config(text="Aucune session ouverte", fg="#aabbcc")

    # ---------------------------------------------------------------- save / load

    def _collect_session_data(self):
        """Sérialise l'état courant (glossaire + infos projet) en dict JSON-compatible."""
        infos = {}
        for attr, key in [
            ("entry_batiment", "batiment"), ("entry_adresse", "adresse"),
            ("entry_proprio",  "proprio"),  ("entry_cadastre", "cadastre"),
            ("entry_date",     "date"),     ("entry_dossier",  "dossier"),
            ("entry_mesurage", "mesurage"),
        ]:
            w = getattr(self, attr, None)
            infos[key] = w.get() if w else ""

        # super_cats : values sont des listes → JSON-compatible
        super_cats_ser = {t: dict(sc) for t, sc in self.super_cats.items()}

        # mappings : DataFrame → liste de {Affectation, cat}
        mappings_ser = {}
        for t, df in self.mappings.items():
            mappings_ser[t] = df[["Affectation", "cat"]].to_dict(orient="records")

        return {
            "version":    SVFILL_VERSION,
            "infos":      infos,
            "super_cats": super_cats_ser,
            "mappings":   mappings_ser,
            "cat_colors": self.cat_colors,
            "loaded_files": self._loaded_files,
        }

    def _apply_session_data(self, data):
        """Restaure infos projet, glossaire et couleurs depuis un dict .svfill."""
        infos = data.get("infos", {})
        for attr, key in [
            ("entry_batiment", "batiment"), ("entry_adresse", "adresse"),
            ("entry_proprio",  "proprio"),  ("entry_cadastre", "cadastre"),
            ("entry_date",     "date"),     ("entry_dossier",  "dossier"),
            ("entry_mesurage", "mesurage"),
        ]:
            w = getattr(self, attr, None)
            if w:
                w.delete(0, "end")
                w.insert(0, infos.get(key, ""))

        self.super_cats = {t: dict(sc) for t, sc in data.get("super_cats", {}).items()}
        self.cat_colors = data.get("cat_colors", {})
        self.saved_glossaries = set()

        # Reconstruire mappings comme DataFrames
        self.mappings = {}
        self.empty_categories = {}
        for t, rows in data.get("mappings", {}).items():
            self.mappings[t] = pd.DataFrame(rows, columns=["Affectation", "cat"])
            self.empty_categories[t] = set()

        # Relire les fichiers Excel silencieusement pour reconstruire df_types
        # (nécessaire pour la génération), SANS toucher aux mappings/super_cats restaurés
        self._loaded_files = data.get("loaded_files", [])
        self._reload_files_silently()

        if self.mappings:
            types = list(self.mappings.keys())
            self.combo_type["values"] = types
            self.combo_type.set(types[0])
            self.load_glossaire_board()

    def session_new(self):
        """Remet l'application à zéro (nouvelle session)."""
        if not messagebox.askyesno("Nouvelle session",
                                   "Abandonner la session courante et repartir de zéro ?"):
            return
        self.clear_data()
        for attr in ("entry_batiment", "entry_adresse", "entry_proprio",
                     "entry_cadastre", "entry_date", "entry_dossier", "entry_mesurage"):
            w = getattr(self, attr, None)
            if w:
                w.delete(0, "end")
        self._session_path = None
        self._update_session_label()

    def session_open(self):
        """Ouvre un fichier .svfill et restaure l'état de la session."""
        path = filedialog.askopenfilename(
            title="Ouvrir une session Rfill",
            filetypes=[("Sessions Rfill", "*.svfill"), ("Tous les fichiers", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lire la session :\n{e}")
            return
        self._apply_session_data(data)
        self._session_path = path
        self._update_session_label()

    def session_save(self):
        """Enregistre la session courante dans le fichier .svfill actif (ou demande le chemin)."""
        if not self._session_path:
            self.session_save_as()
            return
        self._write_session(self._session_path)

    def session_save_as(self):
        """Enregistre la session sous un nouveau fichier .svfill."""
        path = filedialog.asksaveasfilename(
            title="Enregistrer la session",
            defaultextension=".svfill",
            filetypes=[("Sessions Rfill", "*.svfill"), ("Tous les fichiers", "*.*")],
        )
        if not path:
            return
        self._write_session(path)
        self._session_path = path
        self._update_session_label()

    def _reload_files_silently(self):
        """Relit les fichiers Excel enregistrés dans la session pour reconstruire df_types.

        Ne touche pas aux mappings ni aux super_cats déjà restaurés.
        Les fichiers introuvables sont signalés dans le journal d'import.
        """
        if not self._loaded_files:
            return
        dfs = []
        missing = []
        for entry in self._loaded_files:
            fpath = entry.get("path", "")
            if not fpath or not os.path.isfile(fpath):
                missing.append(os.path.basename(fpath) if fpath else "?")
                continue
            try:
                df = extract_info(fpath)
                if not df.empty:
                    dfs.append(df)
            except Exception:
                missing.append(os.path.basename(fpath))

        if missing:
            self._log(self.text_import,
                      f"[WARN] Fichier(s) introuvable(s) à la réouverture : {', '.join(missing)}", "warn")

        if not dfs:
            return

        merged = pd.concat(dfs, ignore_index=True)
        try:
            self.all_df   = merged
            self.df_types = tab_cd_type(merged)
        except Exception as e:
            self._log(self.text_import, f"[ERR]  Rechargement des données : {e}", "err")

    def _write_session(self, path):
        """Sérialise et écrit la session dans path."""
        data = self._collect_session_data()
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'enregistrer la session :\n{e}")
            return
        messagebox.showinfo("OK", f"Session enregistrée :\n{os.path.basename(path)}")

    # ------------------------------------------------------------------ helpers

    def _ts(self):
        """Retourne l'heure courante formatée HH:MM:SS pour les entrées de journal."""
        return datetime.now().strftime("%H:%M:%S")

    def _log(self, widget, msg, tag="info"):
        """Insère [HH:MM:SS] msg dans le widget Text avec le tag de couleur donné."""
        widget.insert("end", f"[{self._ts()}]  {msg}\n", tag)
        widget.see("end")

    def _log_sep(self, widget, label=""):
        """Insère un séparateur visuel (tirets) dans le journal, avec label optionnel."""
        line = f"{'─' * 60}"
        if label:
            line = f"── {label} {'─' * max(0, 57 - len(label))}"
        widget.insert("end", line + "\n", "sep")
        widget.see("end")

    # ================================================================ IMPORT

    def build_import_tab(self):
        """Construit l'onglet Import : boutons de chargement et console de journal."""
        frame = ttk.Frame(self.tab_import, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Importer des fichiers de surfaces",
                  style="Title.TLabel").pack(anchor="w", pady=(0, 4))
        ttk.Label(frame,
                  text="Chargez un ou plusieurs fichiers ; les nouvelles données s'ajoutent aux précédentes.",
                  style="Sub.TLabel").pack(anchor="w", pady=(0, 12))

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(anchor="w")

        ttk.Button(btn_frame, text="AutoCAD  (.xls)",
                   command=lambda: self.load_files("autocad")).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frame, text="GeoGex  (.xlsx)",
                   command=lambda: self.load_files("geogex")).pack(side="left")
        ttk.Button(btn_frame, text="Effacer tout",
                   command=self.clear_data).pack(side="left", padx=(20, 0))

        ttk.Separator(frame, orient="horizontal").pack(fill="x", pady=12)

        # Barre journal
        log_hdr = ttk.Frame(frame)
        log_hdr.pack(fill="x", pady=(0, 4))
        ttk.Label(log_hdr, text="Journal de chargement",
                  font=("Arial", 10, "bold")).pack(side="left")
        ttk.Button(log_hdr, text="Effacer",
                   command=lambda: self.text_import.delete("1.0", "end")
                   ).pack(side="right")

        log_frame = ttk.Frame(frame, relief="sunken", borderwidth=1)
        log_frame.pack(fill="both", expand=True)

        sb = ttk.Scrollbar(log_frame, orient="vertical")
        sb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(log_frame, orient="horizontal")
        hsb.pack(side="bottom", fill="x")

        self.text_import = tk.Text(
            log_frame,
            yscrollcommand=sb.set, xscrollcommand=hsb.set,
            font=("Consolas", 9), bg="#1a1d23", fg="#abb2bf",
            insertbackground="white", relief="flat", bd=6,
            wrap="none",
        )
        self.text_import.pack(fill="both", expand=True)
        sb.config(command=self.text_import.yview)
        hsb.config(command=self.text_import.xview)

        self.text_import.tag_config("ok",   foreground="#98c379")
        self.text_import.tag_config("warn", foreground="#e5c07b")
        self.text_import.tag_config("err",  foreground="#e06c75")
        self.text_import.tag_config("info", foreground="#61afef")
        self.text_import.tag_config("sep",  foreground="#4b5263")

        self._log(self.text_import, f"Session démarrée — {datetime.now().strftime('%d/%m/%Y %H:%M')}", "sep")

    def load_files(self, source):
        """Ouvre le sélecteur de fichiers, parse chaque fichier et accumule les données dans all_df."""
        if source == "autocad":
            filetypes = [("Fichiers AutoCAD XLS", "*.xls"), ("Tous les Excel", "*.xls *.xlsx")]
            title     = "Choisir fichiers AutoCAD (XLS)"
        else:
            filetypes = [("Fichiers GeoGex XLSX", "*.xlsx"), ("Tous les Excel", "*.xls *.xlsx")]
            title     = "Choisir fichiers GeoGex (XLSX)"

        files = filedialog.askopenfilenames(title=title, filetypes=filetypes)
        if not files:
            return

        self._log_sep(self.text_import, f"Chargement {source.upper()}")
        already_loaded = {e["path"] for e in self._loaded_files}

        dfs = []
        errors = []
        for f in files:
            fname = os.path.basename(f)
            if not os.path.isfile(f):
                self._log(self.text_import, f"[ERR]  {fname} — fichier introuvable", "err")
                errors.append(fname)
                continue
            if os.path.getsize(f) == 0:
                self._log(self.text_import, f"[ERR]  {fname} — fichier vide", "err")
                errors.append(fname)
                continue
            try:
                df = extract_info(f)
                if df.empty:
                    raise ValueError("Aucune ligne de données trouvée.")
                dfs.append(df)
                if f not in already_loaded:
                    self._loaded_files.append({"path": f, "source": source})
                self._log(self.text_import, f"[OK]   {fname}  ({len(df)} lignes)", "ok")
            except ValueError as e:
                self._log(self.text_import, f"[ERR]  {fname} — {e}", "err")
                errors.append(fname)
            except Exception as e:
                self._log(self.text_import, f"[ERR]  {fname} — {type(e).__name__}: {e}", "err")
                errors.append(fname)

        if errors and not dfs:
            messagebox.showerror("Aucun fichier valide",
                f"{len(errors)} fichier(s) ont échoué. Voir le journal.")
            return
        if errors:
            messagebox.showwarning("Fichiers partiellement chargés",
                f"{len(errors)} fichier(s) ignoré(s), {len(dfs)} intégré(s).\nVoir le journal.")

        new_df = pd.concat(dfs, ignore_index=True)
        merged = new_df if self.all_df is None else pd.concat(
            [self.all_df, new_df], ignore_index=True
        )
        try:
            df_types = tab_cd_type(merged)
        except Exception as e:
            messagebox.showerror("Erreur de traitement", f"Impossible de traiter les données :\n{e}")
            return
        self.all_df   = merged
        self.df_types = df_types
        self.update_glossaire_tab()

        types_detectes = [d["type_su"].iloc[0] for d in self.df_types]
        self._log(self.text_import,
                  f"[OK]   {len(dfs)} fichier(s) intégré(s)  ·  Types : {types_detectes}", "ok")

    def clear_data(self):
        """Efface toutes les données de session après confirmation utilisateur."""
        if not messagebox.askyesno("Confirmation", "Effacer toutes les données chargées ?"):
            return
        self.all_df = None
        self.df_types = None
        self.mappings = {}
        self.empty_categories = {}
        self._loaded_files = []
        self.combo_type["values"] = []
        self.combo_type.set("")
        for w in self.board.winfo_children():
            w.destroy()
        self._log_sep(self.text_import, "Données effacées")

    # =============================================================== GLOSSAIRE

    def build_glossaire_tab(self):
        """Construit l'onglet Glossaire : panel gauche 'Non classé', board catégories à droite."""
        frame = ttk.Frame(self.tab_glossaire, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Éditeur des catégories",
                  style="Title.TLabel").pack(anchor="w")
        ttk.Label(frame,
                  text="Glissez les affectations entre colonnes  ·  Double-clic sur un titre pour renommer  ·  Clic droit pour déplacer vers une sur-catégorie",
                  style="Sub.TLabel").pack(anchor="w", pady=(2, 10))

        top = ttk.Frame(frame)
        top.pack(fill="x")
        ttk.Label(top, text="Type de surface :").pack(side="left", padx=(0, 6))
        self.combo_type = ttk.Combobox(top, state="readonly", width=22)
        self.combo_type.pack(side="left")
        self.combo_type.bind("<<ComboboxSelected>>", self.load_glossaire_board)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(anchor="w", pady=8)
        ttk.Button(btn_frame, text="+ Ajouter une catégorie",
                   command=self.add_category).pack(side="left", padx=(0, 6))
        ttk.Button(btn_frame, text="Sauv. dans glossaire source ↗",
                   command=self.save_glossaire).pack(side="left")

        # ── Split horizontal : panel gauche "Non classé" + board catégories ──
        split = tk.Frame(frame, bg="#f0f0f0")
        split.pack(fill="both", expand=True)

        # ── Panel gauche ─────────────────────────────────────────────────────
        self._autres_outer = tk.Frame(split, width=195, bg="#e0e0e0")
        self._autres_outer.pack(side="left", fill="y")
        self._autres_outer.pack_propagate(False)

        autres_hdr = tk.Frame(self._autres_outer, bg="#999999", pady=6)
        autres_hdr.pack(fill="x")
        tk.Label(autres_hdr, text="Non classé", font=("Arial", 10, "bold"),
                 bg="#999999", fg="white").pack(padx=8)

        autres_cf = tk.Frame(self._autres_outer, bg="#e8e8e8")
        autres_cf.pack(fill="both", expand=True)

        self.autres_canvas = tk.Canvas(autres_cf, bg="#e8e8e8", highlightthickness=0)
        autres_sb = ttk.Scrollbar(autres_cf, orient="vertical",
                                   command=self.autres_canvas.yview)
        self.autres_canvas.config(yscrollcommand=autres_sb.set)
        autres_sb.pack(side="right", fill="y")
        self.autres_canvas.pack(fill="both", expand=True)

        self.autres_inner = tk.Frame(self.autres_canvas, bg="#e8e8e8")
        self._autres_win = self.autres_canvas.create_window(
            (0, 0), window=self.autres_inner, anchor="nw")
        self.autres_inner.bind(
            "<Configure>",
            lambda e: self.autres_canvas.configure(
                scrollregion=self.autres_canvas.bbox("all")),
        )
        self.autres_canvas.bind(
            "<Configure>",
            lambda e: self.autres_canvas.itemconfig(self._autres_win, width=e.width),
        )

        # Séparateur vertical
        ttk.Separator(split, orient="vertical").pack(side="left", fill="y", padx=2)

        # ── Board catégories (droite) ─────────────────────────────────────────
        canvas_frame = ttk.Frame(split)
        canvas_frame.pack(side="left", fill="both", expand=True)
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)

        self.h_scroll = ttk.Scrollbar(canvas_frame, orient="horizontal")
        self.h_scroll.grid(row=1, column=0, sticky="ew")
        self.v_scroll = ttk.Scrollbar(canvas_frame, orient="vertical")
        self.v_scroll.grid(row=0, column=1, sticky="ns")

        self.canvas = tk.Canvas(
            canvas_frame, bg="#f0f0f0",
            xscrollcommand=self.h_scroll.set,
            yscrollcommand=self.v_scroll.set,
            highlightthickness=0,
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.h_scroll.config(command=self.canvas.xview)
        self.v_scroll.config(command=self.canvas.yview)

        self.board = ttk.Frame(self.canvas)
        self._board_win = self.canvas.create_window((0, 0), window=self.board, anchor="nw")
        self.board.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        # Molette : router vers le bon canvas selon position souris
        self.bind_all("<MouseWheel>", self._route_scroll)

    def _on_canvas_resize(self, event):
        """Adapte la largeur du board à la largeur courante du canvas."""
        self.canvas.itemconfig(self._board_win, width=event.width)

    def _route_scroll(self, event):
        """Route l'événement molette vers le canvas sous le pointeur."""
        wx, wy = event.x_root, event.y_root

        def _over(widget):
            try:
                x, y = widget.winfo_rootx(), widget.winfo_rooty()
                return x <= wx < x + widget.winfo_width() and y <= wy < y + widget.winfo_height()
            except Exception:
                return False

        delta = -1 if event.delta < 0 else 1
        if _over(self.autres_canvas):
            self.autres_canvas.yview_scroll(-delta, "units")
        elif _over(self.canvas):
            self.canvas.yview_scroll(-delta, "units")

    def _autoscroll_canvas(self, event, canvas):
        """Déclenche un auto-scroll quand le curseur approche du bord haut/bas du canvas pendant un drag."""
        try:
            cy = canvas.winfo_rooty()
            ch = canvas.winfo_height()
            y  = event.y_root
            margin = 40
            if y < cy + margin:
                canvas.yview_scroll(-1, "units")
            elif y > cy + ch - margin:
                canvas.yview_scroll(1, "units")
        except Exception:
            pass

    def _on_tab_changed(self, event=None):
        """Affiche un message d'invite dans le board Glossaire si aucune donnée n'est chargée."""
        selected = self.notebook.index(self.notebook.select())
        if selected == 1 and not self.mappings:
            for w in self.board.winfo_children():
                w.destroy()
            tk.Label(
                self.board,
                text="Importez d'abord des fichiers (onglet 1)",
                font=("Arial", 11), fg="#999999", bg="#f0f0f0",
            ).pack(pady=50)

    _DEFAULT_SUPER_CATS = ["Superficies utiles", "Superficies communes", "Superficies annexes"]

    def update_glossaire_tab(self):
        """Recalcule mappings et super_cats depuis df_types, met à jour la combobox de types."""
        if not self.df_types:
            return

        types = []
        # Conserver les mappings déjà édités manuellement (drag-drop, session chargée)
        old_mappings    = self.mappings
        old_super_cats  = self.super_cats
        old_empty_cats  = self.empty_categories
        self.mappings = {}
        self.empty_categories = {}
        self.super_cats = {}

        for df_t in self.df_types:
            t = df_t["type_su"].iloc[0].strip().upper()
            types.append(t)

            if t in old_mappings:
                # Récupérer le mapping existant et y ajouter les nouvelles affectations inconnues
                existing_mapping = old_mappings[t]
                new_mapping = build_affectation_mapping(df_t, t)
                known_aff = set(existing_mapping["Affectation"].tolist())
                truly_new = new_mapping[~new_mapping["Affectation"].isin(known_aff)]
                if not truly_new.empty:
                    mapping_df = pd.concat([existing_mapping, truly_new], ignore_index=True)
                else:
                    mapping_df = existing_mapping.copy()
            else:
                mapping_df = build_affectation_mapping(df_t, t)

            self.mappings[t] = mapping_df
            self.empty_categories[t] = old_empty_cats.get(t, set())

            from glossary_surf import (superficie_names as _sn, predefined_cats as _pc,
                                       denom_surf as _dn, glossary_surf as _gs)
            known_cats = [c for c in mapping_df["cat"].unique() if c != "autres"]

            if t == "SDP":
                # Colonnes SDP figées : toujours les clés du glossaire sdp, ordre fixe
                glo_cats = list(_gs.get("sdp", {}).keys())
                all_cats = list(dict.fromkeys(
                    known_cats + [c for c in glo_cats if c not in known_cats]
                ))
                self.super_cats[t] = old_super_cats.get(t, {"": all_cats})
            else:
                predef = _pc.get(t, [])
                if predef:
                    all_cats = list(dict.fromkeys(
                        known_cats + [c for c in predef if c not in known_cats]
                    ))
                else:
                    glo_key  = _dn.get(t)
                    glo_cats = list(_gs.get(glo_key, {}).keys()) if glo_key else []
                    all_cats = list(dict.fromkeys(
                        known_cats + [c for c in glo_cats if c not in known_cats]
                    ))
                if t in old_super_cats:
                    # Préserver la structure existante, ajouter seulement les nouvelles catégories
                    self.super_cats[t] = old_super_cats[t]
                    already = {c for cats in self.super_cats[t].values() for c in cats}
                    new_cats = [c for c in all_cats if c not in already]
                    if new_cats:
                        first = next(iter(self.super_cats[t]))
                        self.super_cats[t][first].extend(new_cats)
                else:
                    defaults = list(_sn.get(t, self._DEFAULT_SUPER_CATS))
                    self.super_cats[t] = {sc: [] for sc in defaults}
                    if defaults:
                        self.super_cats[t][defaults[0]] = all_cats

            self.saved_glossaries.discard(t)

        self.combo_type["values"] = types

    _COL_COLORS = [
        "#d5e8d4", "#dae8fc", "#fff2cc", "#f8cecc", "#e1d5e7",
        "#fce5cd", "#d0e0e3", "#cfe2f3", "#d9d2e9", "#fde9d9",
    ]

    _SC_HDR_BG  = "#b0c4de"
    _SC_BODY_BG = "#f5f5f5"

    def load_glossaire_board(self, event=None):
        """Reconstruit le panel gauche 'Non classé' et le board des catégories."""
        type_su = self.combo_type.get()
        if not type_su or type_su not in self.mappings:
            return
        mapping_df = self.mappings[type_su]
        sc_dict    = self.super_cats.get(type_su, {})

        # ── Vider les deux zones ──────────────────────────────────────────────
        for w in self.board.winfo_children():
            w.destroy()
        for w in self.autres_inner.winfo_children():
            w.destroy()
        self.columns = {}

        # ── Panel gauche : affectations "Non classé" ──────────────────────────
        autres_items = sorted(
            mapping_df.loc[mapping_df["cat"] == "autres", "Affectation"].tolist()
        )
        if autres_items:
            # La colonne "autres" est la zone entière autres_inner
            self.columns["autres"] = self.autres_inner
            for aff in autres_items:
                self._make_aff_widget(self.autres_inner, aff, "autres")
        else:
            tk.Label(self.autres_inner, text="(aucune)", font=("Arial", 9),
                     fg="#aaaaaa", bg="#e8e8e8").pack(pady=12, padx=8)
            self.columns["autres"] = self.autres_inner

        # ── Board catégories (droite) ─────────────────────────────────────────
        for sc_name, sc_cats in sc_dict.items():
            if sc_name:
                sc_hdr = tk.Frame(self.board, bg=self._SC_HDR_BG, pady=5)
                sc_hdr.pack(fill="x", padx=4, pady=(10, 0))
                sc_lbl = tk.Label(sc_hdr, text=sc_name, font=("Arial", 11, "bold"),
                                  bg=self._SC_HDR_BG, cursor="hand2")
                sc_lbl.pack(side="left", padx=10)
                sc_lbl.bind("<Double-Button-1>",
                            lambda e, sc=sc_name, t=type_su: self.rename_supercat(sc, t))

            sc_body = tk.Frame(self.board, bg=self._SC_BODY_BG, pady=4)
            sc_body.pack(fill="x", padx=4)

            all_sc_cats = list(sc_cats) + [
                c for c in self.empty_categories.get(type_su, set())
                if c not in {x for v in sc_dict.values() for x in v}
                and list(sc_dict.keys()).index(sc_name) == 0
            ]

            if not all_sc_cats:
                tk.Label(sc_body, text="(vide — clic droit sur une catégorie pour la déplacer ici)",
                         font=("Arial", 9), fg="#aaaaaa",
                         bg=self._SC_BODY_BG).pack(side="left", padx=20, pady=8)

            for cat in all_sc_cats:
                colors = self.cat_colors.setdefault(type_su, {})
                if cat not in colors:
                    from glossary_surf import cat_colors_sub as _cc_sub
                    _predef_colors = {"SUB": _cc_sub}
                    predef = _predef_colors.get(type_su, {})
                    if cat in predef:
                        color = predef[cat]
                    else:
                        used = set(colors.values())
                        color = next(
                            c for c in self._COL_COLORS * 2
                            if c not in used
                        ) if len(used) < len(self._COL_COLORS) else self._COL_COLORS[len(colors) % len(self._COL_COLORS)]
                    colors[cat] = color
                color = colors[cat]

                col = tk.Frame(sc_body, bd=1, relief="solid", bg=color, padx=8, pady=8)
                col.pack(side="left", anchor="n", padx=6, pady=6)

                hdr = tk.Frame(col, bg=color)
                hdr.pack(fill="x", pady=(0, 6))

                lbl = tk.Label(hdr, text=cat, font=("Arial", 10, "bold"),
                               bg=color, wraplength=120, justify="center", cursor="fleur")
                lbl.pack(side="left", fill="x", expand=True)
                lbl.bind("<Double-Button-1>", lambda e, c=cat: self.rename_category(c))
                lbl.bind("<Button-3>",
                         lambda e, c=cat, sc=sc_name: self.show_supercat_menu(e, c, sc))
                lbl.bind("<Button-1>",
                         lambda e, c=cat, sc=sc_name: self._cat_drag_start(e, c, sc))
                lbl.bind("<B1-Motion>",       self._cat_drag_do)
                lbl.bind("<ButtonRelease-1>", self._cat_drag_stop)

                del_btn = tk.Label(hdr, text="✕", font=("Arial", 9), bg=color,
                                   fg="#cc0000", cursor="hand2", padx=2)
                del_btn.pack(side="right")
                del_btn.bind("<Button-1>", lambda e, c=cat: self.delete_category(c))

                self.columns[cat] = col

                if cat in mapping_df["cat"].values:
                    for aff in sorted(mapping_df.loc[mapping_df["cat"] == cat, "Affectation"]):
                        self._make_aff_widget(col, aff, cat)

    def _make_aff_widget(self, parent, aff, cat):
        """Crée un label blanc draggable représentant une affectation dans sa colonne de catégorie."""
        lbl = tk.Label(
            parent, text=aff, bg="white", relief="raised",
            padx=5, pady=3, font=("Arial", 9), anchor="w",
            cursor="fleur", wraplength=140,
        )
        lbl.pack(fill="x", pady=2)
        lbl.bind("<Button-1>",        lambda e, a=aff, c=cat: self.start_drag(e, a, c))
        lbl.bind("<B1-Motion>",       self.do_drag)
        lbl.bind("<ButtonRelease-1>", self.stop_drag)

    # ----------------------------------------- sur-catégories (clic droit)

    def show_supercat_menu(self, event, cat, current_sc):
        """Affiche le menu contextuel pour déplacer cat vers une autre sur-catégorie."""
        type_su = self.combo_type.get()
        menu = tk.Menu(self, tearoff=0)
        for sc in self.super_cats.get(type_su, {}):
            if sc != current_sc:
                menu.add_command(
                    label=f"Déplacer vers « {sc} »",
                    command=lambda s=sc, c=cat, cs=current_sc: self.move_cat_to_supercat(c, cs, s),
                )
        menu.tk_popup(event.x_root, event.y_root)

    def move_cat_to_supercat(self, cat, from_sc, to_sc):
        """Déplace cat de from_sc vers to_sc dans super_cats et rafraîchit le board."""
        type_su = self.combo_type.get()
        sc_dict = self.super_cats[type_su]
        if cat in sc_dict.get(from_sc, []):
            sc_dict[from_sc].remove(cat)
        if to_sc in sc_dict and cat not in sc_dict[to_sc]:
            sc_dict[to_sc].append(cat)
        self.load_glossaire_board()

    # ---------------------------------------------------- renommage catégorie

    def rename_category(self, old_name):
        """Renomme une catégorie via dialog et met à jour mappings, empty_categories, super_cats."""
        type_su = self.combo_type.get()
        if not type_su:
            return

        new_name = simpledialog.askstring(
            "Renommer la catégorie",
            f"Nouveau nom pour « {old_name} » :",
            initialvalue=old_name,
        )
        if not new_name or new_name == old_name:
            return

        mapping_df = self.mappings[type_su]
        existing   = set(mapping_df["cat"].unique()) | self.empty_categories.get(type_su, set())
        if new_name in existing:
            messagebox.showerror("Erreur", f"La catégorie « {new_name} » existe déjà.")
            return

        mapping_df["cat"] = mapping_df["cat"].replace(old_name, new_name)
        self.mappings[type_su] = mapping_df

        if old_name in self.empty_categories[type_su]:
            self.empty_categories[type_su].discard(old_name)
            self.empty_categories[type_su].add(new_name)

        for sc_cats in self.super_cats.get(type_su, {}).values():
            if old_name in sc_cats:
                sc_cats[sc_cats.index(old_name)] = new_name

        # Transférer la couleur vers le nouveau nom
        colors = self.cat_colors.get(type_su, {})
        if old_name in colors:
            colors[new_name] = colors.pop(old_name)

        self.load_glossaire_board()

    # ------------------------------------------------------------ drag & drop

    def start_drag(self, event, aff, cat):
        """Démarre le drag d'une affectation : mémorise l'origine et crée le label fantôme orange."""
        self.drag_item       = aff
        self.drag_origin_cat = cat
        self.drag_label      = tk.Label(
            self, text=aff, bg="#f9a825", relief="solid",
            font=("Arial", 9), padx=6, pady=3,
        )
        self.drag_label.place(
            x=event.x_root - self.winfo_rootx() - 50,
            y=event.y_root - self.winfo_rooty() - 20,
        )

    def do_drag(self, event):
        """Met à jour la position du label fantôme et déclenche l'auto-scroll aux bords."""
        if self.drag_label:
            self.drag_label.place(
                x=event.x_root - self.winfo_rootx() - 50,
                y=event.y_root - self.winfo_rooty() - 20,
            )
            self._autoscroll_canvas(event, self.canvas)
            self._autoscroll_canvas(event, self.autres_canvas)

    def stop_drag(self, event):
        """Dépose l'affectation dans la catégorie cible et met à jour le mapping."""
        if not self.drag_label:
            return

        x, y        = event.x_root, event.y_root
        target_cat  = None

        # Vérifier d'abord le panel "Non classé" entier
        try:
            ax, ay = self.autres_canvas.winfo_rootx(), self.autres_canvas.winfo_rooty()
            aw, ah = self.autres_canvas.winfo_width(), self.autres_canvas.winfo_height()
            if ax < x < ax + aw and ay < y < ay + ah:
                target_cat = "autres"
        except Exception:
            pass

        if target_cat is None:
            for cat, frame in self.columns.items():
                if cat == "autres":
                    continue
                fx, fy = frame.winfo_rootx(), frame.winfo_rooty()
                fw, fh = frame.winfo_width(),  frame.winfo_height()
                if fx < x < fx + fw and fy < y < fy + fh:
                    target_cat = cat
                    break

        self.drag_label.destroy()
        self.drag_label = None

        if not target_cat or target_cat == self.drag_origin_cat:
            return

        type_su    = self.combo_type.get()
        mapping_df = self.mappings[type_su]

        self.empty_categories[type_su].discard(target_cat)
        mapping_df.loc[mapping_df["Affectation"] == self.drag_item, "cat"] = target_cat
        self.mappings[type_su] = mapping_df

        self.load_glossaire_board()

    # ---------------------------------------------------- boutons glossaire

    def add_category(self):
        """Crée une catégorie vide via dialog et l'ajoute à la première sur-catégorie."""
        type_su = self.combo_type.get()
        if not type_su:
            messagebox.showwarning("Attention", "Sélectionnez d'abord un type de surface.")
            return

        mapping_df = self.mappings[type_su]
        new_cat    = simpledialog.askstring("Nouvelle catégorie", "Nom de la nouvelle catégorie :")
        if not new_cat:
            return

        existing = set(mapping_df["cat"].unique()) | self.empty_categories[type_su]
        if new_cat in existing:
            messagebox.showerror("Erreur", "Cette catégorie existe déjà.")
            return

        self.empty_categories[type_su].add(new_cat)
        first_sc = next(iter(self.super_cats.get(type_su, {})), None)
        if first_sc:
            self.super_cats[type_su][first_sc].append(new_cat)
        self.load_glossaire_board()

    def delete_category(self, cat):
        """Supprime une catégorie après confirmation et reporte ses affectations vers 'autres'."""
        type_su    = self.combo_type.get()
        if not type_su:
            return
        mapping_df = self.mappings[type_su]
        count      = int((mapping_df["cat"] == cat).sum())
        msg = (
            f"Supprimer « {cat} » ?\n{count} affectation(s) seront déplacées vers « autres »."
            if count else f"Supprimer la catégorie vide « {cat} » ?"
        )
        if not messagebox.askyesno("Supprimer la catégorie", msg):
            return
        mapping_df.loc[mapping_df["cat"] == cat, "cat"] = "autres"
        self.mappings[type_su] = mapping_df
        self.empty_categories[type_su].discard(cat)
        for sc_cats in self.super_cats.get(type_su, {}).values():
            if cat in sc_cats:
                sc_cats.remove(cat)
        self.cat_colors.get(type_su, {}).pop(cat, None)
        self.load_glossaire_board()

    def save_glossaire(self):
        """Persiste les affectations et la structure des catégories dans glossary_surf.py."""
        type_su = self.combo_type.get()
        if not type_su:
            return
        mapping_df = self.mappings[type_su]
        updated    = update_glossary(mapping_df, type_su, self.super_cats.get(type_su, {}))
        self.saved_glossaries.add(type_su)
        msg = (f"Glossaire mis à jour pour {type_su} — catégories et affectations sauvegardées."
               if updated else f"Glossaire sauvegardé pour {type_su} (aucune nouveauté).")
        messagebox.showinfo("OK", msg)

    # ------------------------------------------------ renommage superficie (super-cat)

    def rename_supercat(self, old_sc, type_su):
        """Renomme une sur-catégorie en reconstruisant super_cats avec la nouvelle clé à la même position."""
        new_sc = simpledialog.askstring(
            "Renommer la superficie",
            f"Nouveau nom pour « {old_sc} » :",
            initialvalue=old_sc,
        )
        if not new_sc or new_sc == old_sc:
            return
        sc_dict = self.super_cats.get(type_su, {})
        if new_sc in sc_dict:
            messagebox.showerror("Erreur", f"La superficie « {new_sc} » existe déjà.")
            return
        new_dict = {}
        for k, v in sc_dict.items():
            new_dict[new_sc if k == old_sc else k] = v
        self.super_cats[type_su] = new_dict
        self.load_glossaire_board()

    # ----------------------------------------- drag & drop réordonnancement catégories

    def _cat_drag_start(self, event, cat, sc):
        """Initialise le contexte de drag pour réordonner une colonne de catégorie."""
        self._cat_drag = {
            "cat": cat, "sc": sc,
            "start_x": event.x_root, "start_y": event.y_root,
            "active": False,
        }

    def _cat_drag_do(self, event):
        """Active le mode drag catégorie après 6 px de mouvement et déplace le fantôme."""
        if not self._cat_drag:
            return
        dx = abs(event.x_root - self._cat_drag["start_x"])
        dy = abs(event.y_root - self._cat_drag["start_y"])
        if not self._cat_drag["active"] and (dx > 6 or dy > 6):
            self._cat_drag["active"] = True
            self._cat_drag_label = tk.Label(
                self, text=self._cat_drag["cat"],
                bg="#ff9800", fg="white", relief="solid",
                font=("Arial", 9, "bold"), padx=6, pady=3,
            )
        if self._cat_drag["active"] and self._cat_drag_label:
            self._cat_drag_label.place(
                x=event.x_root - self.winfo_rootx() - 50,
                y=event.y_root - self.winfo_rooty() - 20,
            )
            self._autoscroll_canvas(event, self.canvas)

    def _cat_drag_stop(self, event):
        """Insère la catégorie source avant la cible dans super_cats et rafraîchit le board."""
        if not self._cat_drag or not self._cat_drag.get("active"):
            self._cat_drag = None
            return
        if self._cat_drag_label:
            self._cat_drag_label.destroy()
            self._cat_drag_label = None

        source_cat = self._cat_drag["cat"]
        source_sc  = self._cat_drag["sc"]
        self._cat_drag = None

        x, y = event.x_root, event.y_root
        target_cat = None
        for cat, frame in self.columns.items():
            fx, fy = frame.winfo_rootx(), frame.winfo_rooty()
            fw, fh = frame.winfo_width(), frame.winfo_height()
            if fx < x < fx + fw and fy < y < fy + fh:
                target_cat = cat
                break

        if not target_cat or target_cat == source_cat:
            return

        type_su = self.combo_type.get()
        sc_dict = self.super_cats[type_su]

        # Trouver la super-cat cible
        target_sc = next(
            (sc for sc, cats in sc_dict.items() if target_cat in cats), None
        )
        if target_sc is None:
            return

        # Retirer la cat source de son emplacement
        if source_cat in sc_dict.get(source_sc, []):
            sc_dict[source_sc].remove(source_cat)

        # Insérer avant la cat cible dans sa super-cat
        t_list = sc_dict[target_sc]
        idx = t_list.index(target_cat) if target_cat in t_list else len(t_list)
        t_list.insert(idx, source_cat)

        self.load_glossaire_board()

    # =============================================================== GÉNÉRER

    def build_generer_tab(self):
        """Construit l'onglet Générer : formulaire projet, bouton d'export, console."""
        frame = ttk.Frame(self.tab_generer, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Informations du projet",
                  style="Title.TLabel").pack(anchor="w", pady=(0, 12))

        form = ttk.Frame(frame)
        form.pack(anchor="w")

        fields = [
            ("Bâtiment",     "entry_batiment"),
            ("Adresse",      "entry_adresse"),
            ("Propriétaire", "entry_proprio"),
            ("Cadastre",     "entry_cadastre"),
            ("Date",         "entry_date"),
            ("Dossier",      "entry_dossier"),
            ("Mesurage",     "entry_mesurage"),
        ]

        for row, (label, attr) in enumerate(fields):
            ttk.Label(form, text=label + " :").grid(
                row=row, column=0, sticky="e", padx=(0, 10), pady=3
            )
            entry = ttk.Entry(form, width=50)
            entry.grid(row=row, column=1, sticky="w", pady=3)
            setattr(self, attr, entry)

        ttk.Separator(frame, orient="horizontal").pack(fill="x", pady=15)

        ttk.Label(frame, text="Export Excel + HTML",
                  style="Title.TLabel").pack(anchor="w", pady=(0, 8))
        ttk.Label(frame,
                  text="Deux fichiers sont générés côte à côte : un classeur Excel (.xlsx, un onglet par type) "
                       "et un rapport HTML imprimable A4 paysage (un onglet par type, logos et tampon intégrés).",
                  style="Sub.TLabel", wraplength=900, justify="left").pack(anchor="w", pady=(0, 8))

        action_row = ttk.Frame(frame)
        action_row.pack(anchor="w", pady=(0, 8))

        ttk.Button(
            action_row, text="  Générer Excel + HTML  ",
            style="Primary.TButton",
            command=self.generate_final_tables,
        ).pack(side="left", padx=(0, 16))

        ttk.Button(
            action_row, text="  HTML depuis Excel…  ",
            command=self.generate_html_from_excel,
        ).pack(side="left", padx=(0, 16))

        self.var_open_after = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            action_row, text="Ouvrir le fichier après enregistrement",
            variable=self.var_open_after,
        ).pack(side="left")

        ttk.Separator(frame, orient="horizontal").pack(fill="x", pady=12)

        log_hdr = ttk.Frame(frame)
        log_hdr.pack(fill="x", pady=(0, 4))
        ttk.Label(log_hdr, text="Journal de génération",
                  font=("Arial", 10, "bold")).pack(side="left")
        ttk.Button(log_hdr, text="Effacer",
                   command=lambda: self.text_generer.delete("1.0", "end")
                   ).pack(side="right")

        log_frame = ttk.Frame(frame, relief="sunken", borderwidth=1)
        log_frame.pack(fill="both", expand=True)

        sb = ttk.Scrollbar(log_frame, orient="vertical")
        sb.pack(side="right", fill="y")

        self.text_generer = tk.Text(
            log_frame, yscrollcommand=sb.set,
            font=("Consolas", 9), bg="#1a1d23", fg="#abb2bf",
            relief="flat", bd=6, wrap="none",
        )
        self.text_generer.pack(fill="both", expand=True)
        sb.config(command=self.text_generer.yview)

        self.text_generer.tag_config("ok",   foreground="#98c379")
        self.text_generer.tag_config("warn", foreground="#e5c07b")
        self.text_generer.tag_config("err",  foreground="#e06c75")
        self.text_generer.tag_config("info", foreground="#61afef")
        self.text_generer.tag_config("sep",  foreground="#4b5263")

    def generate_final_tables(self):
        """Sauvegarde les glossaires, calcule les tableaux et exporte le fichier Excel + HTML."""
        if not self.mappings:
            messagebox.showwarning("Attention", "Chargez d'abord des fichiers XLS.")
            return

        unsaved = [t for t in self.mappings if t not in self.saved_glossaries]
        if unsaved:
            for t in unsaved:
                try:
                    update_glossary(self.mappings[t], t, self.super_cats.get(t, {}))
                    self.saved_glossaries.add(t)
                except Exception as e:
                    self._log(self.text_generer,
                              f"[WARN] Sauvegarde auto glossaire {t} échouée : {e}", "warn")

        output_path = filedialog.asksaveasfilename(
            title="Sauvegarder le fichier Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not output_path:
            return

        self._log_sep(self.text_generer, "Génération")

        infos = {
            "batiment": self.entry_batiment.get(),
            "adresse":  self.entry_adresse.get(),
            "proprio":  self.entry_proprio.get(),
            "cadastre": self.entry_cadastre.get(),
            "date":     self.entry_date.get(),
            "dossier":  self.entry_dossier.get(),
            "mesurage": self.entry_mesurage.get(),
        }

        all_tables = {}

        for type_su in self.mappings:
            mapping_df = self.mappings[type_su]
            df_t = next(
                (d for d in self.df_types if d["type_su"].iloc[0] == type_su), None
            )
            if df_t is None:
                continue

            try:
                df_tcd, _ = TCD2Tab(df_t, type_su, mapping_df=mapping_df)
                tables    = Tab_output(df_tcd, infos,
                                       super_cat_map={type_su: self.super_cats.get(type_su, {})})
                all_tables.update(tables)
                self._log(self.text_generer, f"[OK]   {type_su} — tableau calculé", "ok")
            except Exception as e:
                self._log(self.text_generer, f"[ERR]  {type_su} → {e}", "err")

        if not all_tables:
            messagebox.showerror("Erreur", "Aucun tableau généré.")
            return

        try:
            export_tables_to_excel(all_tables, output_path)
            self._log(self.text_generer, f"[OK]   Fichier sauvegardé : {output_path}", "ok")
        except Exception as e:
            self._log(self.text_generer, f"[ERR]  Export Excel → {e}", "err")
            messagebox.showerror("Erreur export", str(e))
            return

        html_path = os.path.splitext(output_path)[0] + ".html"
        try:
            export_tables_to_html(all_tables, infos, html_path)
            self._log(self.text_generer, f"[OK]   Rapport HTML : {html_path}", "ok")
        except Exception as e:
            self._log(self.text_generer, f"[WARN] Export HTML → {e}", "warn")
            html_path = None

        msg = f"Fichier Excel généré :\n{output_path}"
        if html_path:
            msg += f"\n\nRapport HTML (impression A4 paysage) :\n{html_path}"
        messagebox.showinfo("Succès", msg)
        if self.var_open_after.get():
            os.startfile(output_path)

    def generate_html_from_excel(self):
        """Relit un Excel Rfill existant et régénère le rapport HTML à partir des infos du formulaire."""
        xlsx_path = filedialog.askopenfilename(
            title="Choisir un fichier Excel Rfill",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not xlsx_path:
            return

        html_path = filedialog.asksaveasfilename(
            title="Sauvegarder le rapport HTML",
            defaultextension=".html",
            filetypes=[("HTML", "*.html")],
            initialfile=os.path.splitext(os.path.basename(xlsx_path))[0] + ".html",
            initialdir=os.path.dirname(xlsx_path),
        )
        if not html_path:
            return

        self._log_sep(self.text_generer, "HTML depuis Excel")

        infos = {
            "batiment": self.entry_batiment.get(),
            "adresse":  self.entry_adresse.get(),
            "proprio":  self.entry_proprio.get(),
            "cadastre": self.entry_cadastre.get(),
            "date":     self.entry_date.get(),
            "dossier":  self.entry_dossier.get(),
            "mesurage": self.entry_mesurage.get(),
        }

        try:
            output_tables = html_from_excel(xlsx_path, infos)
            self._log(self.text_generer,
                      f"[OK]   {len(output_tables)} feuille(s) lue(s) : {', '.join(output_tables)}", "ok")
        except Exception as e:
            self._log(self.text_generer, f"[ERR]  Lecture Excel → {e}", "err")
            messagebox.showerror("Erreur lecture", str(e))
            return

        try:
            export_tables_to_html(output_tables, infos, html_path)
            self._log(self.text_generer, f"[OK]   Rapport HTML : {html_path}", "ok")
        except Exception as e:
            self._log(self.text_generer, f"[ERR]  Export HTML → {e}", "err")
            messagebox.showerror("Erreur export HTML", str(e))
            return

        messagebox.showinfo("Succès", f"Rapport HTML généré :\n{html_path}")
        if self.var_open_after.get():
            os.startfile(html_path)

    # =============================================================== AIDE

    def build_aide_tab(self):
        """Construit l'onglet Aide : guide complet scrollable."""
        BG = "#f7f9fc"
        DARK = "#1a2a4a"
        MID  = "#3d5a80"

        outer = ttk.Frame(self.tab_aide)
        outer.pack(fill="both", expand=True)

        sb = ttk.Scrollbar(outer, orient="vertical")
        sb.pack(side="right", fill="y")
        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0, yscrollcommand=sb.set)
        canvas.pack(fill="both", expand=True)
        sb.config(command=canvas.yview)

        inner = tk.Frame(canvas, bg=BG)
        win = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))

        PX = 48  # marge gauche/droite

        def titre(text):
            tk.Frame(inner, height=2, bg=MID).pack(fill="x", padx=PX, pady=(28, 0))
            tk.Label(inner, text=text, font=("Arial", 13, "bold"),
                     bg=BG, fg=DARK, anchor="w").pack(fill="x", padx=PX, pady=(6, 2))

        def para(text):
            tk.Label(inner, text=text, font=("Arial", 10),
                     bg=BG, fg="#333", anchor="w", justify="left",
                     wraplength=860).pack(fill="x", padx=PX + 8, pady=(2, 0))

        def bullet(items):
            for item in items:
                row = tk.Frame(inner, bg=BG)
                row.pack(fill="x", padx=PX + 8, pady=1)
                tk.Label(row, text="•", font=("Arial", 10, "bold"),
                         bg=BG, fg=MID, width=2, anchor="w").pack(side="left")
                tk.Label(row, text=item, font=("Arial", 10),
                         bg=BG, fg="#333", anchor="w", justify="left",
                         wraplength=840).pack(side="left", fill="x", expand=True)

        def btn_doc(label, desc):
            row = tk.Frame(inner, bg="#eef4fb", bd=1, relief="flat")
            row.pack(fill="x", padx=PX + 8, pady=3)
            tk.Label(row, text=f"  {label}  ", font=("Arial", 9, "bold"),
                     bg=MID, fg="white", padx=6, pady=3).pack(side="left")
            tk.Label(row, text=desc, font=("Arial", 9),
                     bg="#eef4fb", fg="#444", padx=10, pady=3,
                     anchor="w", justify="left", wraplength=760).pack(side="left", fill="x")

        # ── Entête ────────────────────────────────────────────────────────────
        tk.Frame(inner, height=10, bg=BG).pack()
        tk.Label(inner, text="Guide d'utilisation — Rfill",
                 font=("Arial", 18, "bold"), bg=BG, fg=DARK).pack(anchor="w", padx=PX)
        tk.Label(inner,
                 text="Rfill permet d'importer des fichiers de surfaces AutoCAD ou GeoGex, "
                      "de catégoriser les affectations, puis de générer des tableaux Excel normés "
                      "et un rapport HTML imprimable au format A4 paysage.",
                 font=("Arial", 10, "italic"), bg=BG, fg="#555",
                 anchor="w", justify="left", wraplength=860).pack(anchor="w", padx=PX, pady=(4, 0))

        # ── Session (barre du haut) ───────────────────────────────────────────
        titre("Barre de session  (en haut de la fenêtre)")
        para("La barre bleue en haut permet de gérer votre session de travail. "
             "Une session (.svfill) sauvegarde tout votre travail : fichiers importés, "
             "catégorisations, structure des colonnes et informations du projet.")
        btn_doc("Nouveau",           "Repart de zéro — efface toutes les données et le glossaire en cours.")
        btn_doc("Ouvrir…",           "Charge un fichier .svfill existant : relit les fichiers Excel, "
                                     "restaure l'onglet Glossaire exactement tel qu'il était et remplit "
                                     "le formulaire Générer.")
        btn_doc("Enregistrer",       "Sauvegarde la session courante dans le fichier .svfill actif. "
                                     "Ne modifie PAS le glossaire source.")
        btn_doc("Enregistrer sous…", "Sauvegarde sous un nouveau nom .svfill.")

        # ── Onglet 1 : Importer ───────────────────────────────────────────────
        titre("Onglet 1 — Importer")
        para("Chargez un ou plusieurs fichiers d'étiquettes exportés depuis AutoCAD ou GeoGex. "
             "Les données de chaque fichier s'accumulent dans la session courante.")
        btn_doc("AutoCAD (.xls)",  "Ouvre un sélecteur de fichiers .xls (export AutoCAD natif). "
                                   "Le type de surface est lu dans la colonne Calque "
                                   "(ex. « SUB Contours » → type SUB).")
        btn_doc("GeoGex (.xlsx)",  "Ouvre un sélecteur de fichiers .xlsx (export GeoGex). "
                                   "Même structure de colonnes attendue.")
        btn_doc("Effacer tout",    "Supprime toutes les données chargées et remet le glossaire à zéro. "
                                   "Une confirmation est demandée.")
        para("\nLe journal de chargement (fond sombre) affiche en temps réel le résultat "
             "de chaque fichier avec horodatage. Les erreurs apparaissent en rouge, "
             "les succès en vert.")
        bullet([
            "Chaque étage doit être un chiffre ou un nombre dans la colonne Étage "
            "(ex. -1 pour le 1er sous-sol, 0 pour le rez-de-chaussée, 1 pour le 1er étage, "
            "-0.5 pour rez-de-jardin, 0.5 pour entresol…). "
            "Rfill convertit automatiquement ces valeurs en libellés lisibles dans les tableaux.",
            "Plusieurs fichiers peuvent être chargés en une seule fois (Ctrl+clic).",
            "Les fichiers déjà chargés ne sont pas doublonnés à la réouverture d'une session.",
        ])

        # ── Onglet 2 : Glossaire ──────────────────────────────────────────────
        titre("Onglet 2 — Glossaire")
        para("L'onglet Glossaire permet de contrôler comment chaque affectation du fichier "
             "est rattachée à une catégorie de surface, et comment ces catégories sont "
             "regroupées en colonnes dans le tableau final.")
        para("\nSélectionnez le type de surface à éditer dans la liste déroulante (SUB, SU, SHO…). "
             "Chaque type a son propre classement indépendant.")
        bullet([
            "Colonne « Non classé » (gauche) : affectations non reconnues par le glossaire. "
            "Faites-les glisser vers la bonne catégorie.",
            "Glisser-déposer une affectation (étiquette blanche) → dépose dans la catégorie cible.",
            "Glisser-déposer un titre de catégorie → réordonne les catégories au sein d'une super-catégorie.",
            "Double-clic sur un titre de catégorie → renomme la catégorie.",
            "Double-clic sur un titre de super-catégorie (bandeau bleu) → renomme la super-catégorie.",
            "Clic droit sur un titre de catégorie → déplace la catégorie vers une autre super-catégorie.",
            "✕ (croix rouge) sur une catégorie → supprime la catégorie "
            "(ses affectations retournent dans « Non classé »). "
            "Une catégorie supprimée n'apparaîtra pas dans le tableau final.",
            "+ Ajouter une catégorie → crée une catégorie vide dans la première super-catégorie.",
        ])
        para("\nToutes les catégories présentes dans le glossaire apparaissent dans le tableau final, "
             "même si elles sont vides (valeur 0). Seules les catégories supprimées disparaissent.")
        btn_doc("Sauv. dans glossaire source ↗",
                "Enregistre les nouvelles affectations et la structure des catégories dans le "
                "glossaire source (glossary_surf.py). Ces apprentissages seront réutilisés "
                "automatiquement lors des prochains imports. "
                "ATTENTION : cela modifie le fichier source de manière permanente.")

        # ── Onglet 3 : Générer ────────────────────────────────────────────────
        titre("Onglet 3 — Générer")
        para("Renseignez les informations du projet, puis lancez la génération.")
        bullet([
            "Bâtiment : référence ou nom du bâtiment.",
            "Adresse : adresse complète.",
            "Propriétaire : nom du propriétaire ou du client.",
            "Cadastre : référence cadastrale.",
            "Date : date d'édition (ex. 01/06/2026).",
            "Dossier : numéro de dossier (ex. 2026.T005).",
            "Mesurage : précisez le mois et l'année du mesurage (ex. mai 2026). "
            "Cette valeur est insérée automatiquement dans la note de bas de page des tableaux.",
        ])
        btn_doc("Générer Excel + HTML",
                "Crée deux fichiers côte à côte : un classeur .xlsx (un onglet par type de surface) "
                "et un rapport .html imprimable A4 paysage avec logos et tampon intégrés. "
                "Les étages sont affichés avec leurs libellés complets "
                "(Rez-de-chaussée, 1er Étage, 1er Sous-sol…).")
        btn_doc("HTML depuis Excel…",
                "Relit un fichier .xlsx déjà généré par Rfill et régénère uniquement le rapport HTML, "
                "en utilisant les informations saisies dans le formulaire ci-dessus.")
        btn_doc("☑ Ouvrir après enregistrement",
                "Si coché, ouvre automatiquement le fichier Excel dans votre application par défaut "
                "dès que la génération est terminée.")
        para("\nLe journal de génération affiche l'avancement et les éventuelles erreurs.")

        # ── Codes de surface ──────────────────────────────────────────────────
        titre("Codes de surface reconnus")
        codes = [
            ("SUB",  "Superficie Utile Brute"),
            ("SU",   "Superficie Utile"),
            ("SUBL", "Superficie Utile Brute Locative"),
            ("SUN",  "Superficie Utile Nette"),
            ("SHO",  "Superficie Hors Œuvre"),
            ("SDP",  "Superficie De Plancher"),
            ("GLA",  "Superficie Globale"),
            ("TAX",  "Surfaces Réelles — Art. 324 M ou Z Annexe III CGI"),
            ("TSB",  "Tableau Surfaces Brutes"),
        ]
        for code, label in codes:
            row = tk.Frame(inner, bg=BG)
            row.pack(fill="x", padx=PX + 8, pady=1)
            tk.Label(row, text=code, font=("Arial", 10, "bold"),
                     bg=BG, fg=MID, width=6, anchor="w").pack(side="left")
            tk.Label(row, text=label, font=("Arial", 10),
                     bg=BG, fg="#333", anchor="w").pack(side="left")

        # ── Conseils ──────────────────────────────────────────────────────────
        titre("Conseils d'utilisation")
        bullet([
            "Sauvegardez régulièrement votre session (.svfill) — elle contient tout votre travail "
            "de catégorisation et peut être repartagée à un collègue.",
            "N'utilisez « Sauv. dans glossaire source ↗ » que lorsque vous êtes sûr de vos "
            "affectations : cela modifie le glossaire commun à tous les projets.",
            "Pour un nouveau projet avec le même type de bâtiment, ouvrez une ancienne session "
            "similaire, chargez les nouveaux fichiers : vos catégorisations passées seront "
            "automatiquement réappliquées.",
            "Le champ Étage dans vos fichiers d'étiquettes doit être un nombre "
            "(pas de texte comme « RdC » ou « SS »). Rfill se charge de la conversion en libellé.",
        ])

        tk.Frame(inner, height=40, bg=BG).pack()

    # =============================================================== À PROPOS

    def build_apropos_tab(self):
        """Construit l'onglet À propos."""
        BG     = "#f7f9fc"
        DARK   = "#1a2a4a"
        CARD   = "#eef2f7"
        ACCENT = "#b0c4de"
        MID    = "#3d5a80"

        canvas = tk.Canvas(self.tab_apropos, bg=BG, highlightthickness=0)
        canvas.pack(fill="both", expand=True)

        center = tk.Frame(canvas, bg=BG)

        def _place_center(event=None):
            canvas.update_idletasks()
            x = max(0, (canvas.winfo_width()  - center.winfo_reqwidth())  // 2)
            y = max(0, (canvas.winfo_height() - center.winfo_reqheight()) // 2)
            canvas.coords(win, x, y)

        win = canvas.create_window(0, 0, window=center, anchor="nw")
        canvas.bind("<Configure>", _place_center)

        # ── Titre ─────────────────────────────────────────────────────────────
        tk.Label(center, text=APP_NAME,
                 font=("Arial", 52, "bold"), fg=DARK, bg=BG).pack()
        tk.Label(center, text=f"by {APP_COMPANY}",
                 font=("Arial", 13, "italic"), fg=MID, bg=BG).pack()
        tk.Frame(center, height=3, bg=ACCENT, width=340).pack(pady=(6, 10))
        tk.Label(center, text="Analyse et génération de tableaux de surfaces immobilières",
                 font=("Arial", 10, "italic"), fg="#666", bg=BG).pack()

        tk.Frame(center, height=20, bg=BG).pack()

        # ── Carte infos ───────────────────────────────────────────────────────
        card = tk.Frame(center, bg=CARD, bd=1, relief="solid", padx=40, pady=22)
        card.pack()

        def sep():
            tk.Frame(card, height=1, bg=ACCENT).pack(fill="x", pady=(10, 8))

        def row(label, value, value_font=None, value_fg="#222"):
            r = tk.Frame(card, bg=CARD)
            r.pack(fill="x", pady=2)
            tk.Label(r, text=label, font=("Arial", 9, "bold"),
                     fg="#777", bg=CARD, width=20, anchor="e").pack(side="left")
            tk.Label(r, text=value, font=value_font or ("Arial", 10),
                     fg=value_fg, bg=CARD, anchor="w").pack(side="left", padx=(10, 0))

        row("Version :",      APP_VERSION)
        row("Build :",        APP_BUILD)
        row("Mise à jour :",  APP_DATE)

        sep()

        row("Éditeur :",      APP_COMPANY,
            value_font=("Arial", 11, "bold"), value_fg=DARK)
        row("Développé par :", "J. FAGUET",
            value_font=("Arial", 11, "bold"), value_fg=MID)

        sep()

        tk.Label(card, text=APP_DESC,
                 font=("Arial", 9), fg="#555", bg=CARD, justify="center").pack()

        tk.Frame(center, height=22, bg=BG).pack()

        # ── Bloc légal ────────────────────────────────────────────────────────
        legal_frame = tk.Frame(center, bg=BG)
        legal_frame.pack(padx=20)

        tk.Label(legal_frame, text=APP_COPY,
                 font=("Arial", 9, "bold"), fg="#444", bg=BG).pack()
        tk.Frame(legal_frame, height=5, bg=BG).pack()
        tk.Label(legal_frame, text=APP_LEGAL,
                 font=("Arial", 8), fg="#999", bg=BG,
                 justify="center", wraplength=540).pack()
        tk.Frame(legal_frame, height=10, bg=BG).pack()
        tk.Label(legal_frame,
                 text="Logiciel développé par Jules FAGUET pour RTaxes.",
                 font=("Arial", 9, "italic"), fg="#666", bg=BG).pack()


if __name__ == "__main__":
    app = SurfaceApp()
    app.mainloop()
