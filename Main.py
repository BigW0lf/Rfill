import os
from datetime import datetime

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from process_surf import (
    extract_info, tab_cd_type, build_affectation_mapping,
    TCD2Tab, Tab_output, export_tables_to_excel, export_tables_to_html,
    update_glossary,
)

APP_NAME    = "Rfill"
APP_TITLE   = "Rfill — Analyse des Surfaces"
APP_VERSION = "1.0.0"
APP_BUILD   = "2026.04.27"
APP_DATE    = "27 avril 2026"
APP_AUTHOR  = "Jules FAGUET"
APP_DESC    = (
    "Importation et catégorisation de surfaces AutoCAD/GeoGex,\n"
    "génération de tableaux Excel normés (SUB, SU, SHO, SDP…)."
)
APP_COPY    = "© 2026 Rfill — Tous droits réservés"


class SurfaceApp(tk.Tk):
    def __init__(self):
        """Initialise la fenêtre principale, les styles ttk et les 5 onglets."""
        super().__init__()

        self.title(APP_TITLE)
        self.geometry("1200x780")
        self.configure(bg="#f0f0f0")

        self._setup_styles()

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
        self.combo_type["values"] = []
        self.combo_type.set("")
        for w in self.board.winfo_children():
            w.destroy()
        self._log_sep(self.text_import, "Données effacées")

    # =============================================================== GLOSSAIRE

    def build_glossaire_tab(self):
        """Construit l'onglet Glossaire : combobox, boutons, canvas scrollable et board."""
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
        ttk.Button(btn_frame, text="Sauvegarder",
                   command=self.save_glossaire).pack(side="left")

        canvas_frame = ttk.Frame(frame)
        canvas_frame.pack(fill="both", expand=True)
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

    def _on_canvas_resize(self, event):
        """Adapte la largeur du board à la largeur courante du canvas."""
        self.canvas.itemconfig(self._board_win, width=event.width)

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
        self.mappings = {}
        self.empty_categories = {}

        for df_t in self.df_types:
            t = df_t["type_su"].iloc[0]
            types.append(t)
            mapping_df = build_affectation_mapping(df_t, t)
            self.mappings[t] = mapping_df
            self.empty_categories[t] = set()

            from glossary_surf import superficie_names as _sn, predefined_cats as _pc
            known_cats  = [c for c in mapping_df["cat"].unique() if c != "autres"]
            predef      = _pc.get(t, [])
            # Catégories à afficher : issues des données + prédéfinies (ordre: données en premier)
            all_cats = list(dict.fromkeys(known_cats + [c for c in predef if c not in known_cats]))

            if t not in self.super_cats:
                defaults = list(_sn.get(t, self._DEFAULT_SUPER_CATS))
                self.super_cats[t] = {sc: [] for sc in defaults}
                if defaults:
                    self.super_cats[t][defaults[0]] = all_cats
            else:
                already = {c for cats in self.super_cats[t].values() for c in cats}
                new_cats = [c for c in all_cats if c not in already]
                if new_cats:
                    first = next(iter(self.super_cats[t]))
                    self.super_cats[t][first].extend(new_cats)

            self.saved_glossaries.discard(t)

        self.combo_type["values"] = types

    _COL_COLORS = [
        "#d5e8d4", "#dae8fc", "#fff2cc", "#f8cecc", "#e1d5e7",
        "#fce5cd", "#d0e0e3", "#cfe2f3", "#d9d2e9", "#fde9d9",
    ]

    _SC_HDR_BG  = "#b0c4de"
    _SC_BODY_BG = "#f5f5f5"

    def load_glossaire_board(self, event=None):
        """Reconstruit entièrement le board des catégories pour le type de surface sélectionné."""
        type_su = self.combo_type.get()
        if not type_su or type_su not in self.mappings:
            return
        mapping_df = self.mappings[type_su]
        sc_dict    = self.super_cats.get(type_su, {})

        for w in self.board.winfo_children():
            w.destroy()
        self.columns = {}

        for sc_name, sc_cats in sc_dict.items():
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
                color = self._COL_COLORS[sum(ord(c) for c in cat) % len(self._COL_COLORS)]

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

        autres_items = mapping_df.loc[mapping_df["cat"] == "autres", "Affectation"].tolist()
        if autres_items:
            sc_hdr = tk.Frame(self.board, bg="#cccccc", pady=5)
            sc_hdr.pack(fill="x", padx=4, pady=(10, 0))
            tk.Label(sc_hdr, text="Non classé", font=("Arial", 11, "bold"),
                     bg="#cccccc").pack(side="left", padx=10)

            sc_body = tk.Frame(self.board, bg=self._SC_BODY_BG, pady=4)
            sc_body.pack(fill="x", padx=4)

            col = tk.Frame(sc_body, bd=1, relief="solid", bg="#e8e8e8", padx=8, pady=8)
            col.pack(side="left", anchor="n", padx=6, pady=6)
            tk.Label(col, text="autres", font=("Arial", 10, "bold"), bg="#e8e8e8").pack(fill="x", pady=(0, 6))
            self.columns["autres"] = col
            for aff in sorted(autres_items):
                self._make_aff_widget(col, aff, "autres")

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
        """Met à jour la position du label fantôme lors du déplacement."""
        if self.drag_label:
            self.drag_label.place(
                x=event.x_root - self.winfo_rootx() - 50,
                y=event.y_root - self.winfo_rooty() - 20,
            )

    def stop_drag(self, event):
        """Dépose l'affectation dans la catégorie cible et met à jour le mapping."""
        if not self.drag_label:
            return

        x, y        = event.x_root, event.y_root
        target_cat  = None

        for cat, frame in self.columns.items():
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
        self.load_glossaire_board()

    def save_glossaire(self):
        """Persiste les nouvelles affectations classifiées dans glossary_surf.py via update_glossary."""
        type_su = self.combo_type.get()
        if not type_su:
            return
        mapping_df = self.mappings[type_su]
        updated    = update_glossary(mapping_df, type_su)
        self.saved_glossaries.add(type_su)
        msg = (f"Glossaire mis à jour pour {type_su} — nouvelles affectations ajoutées."
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
                    update_glossary(self.mappings[t], t)
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

    # =============================================================== AIDE

    def build_aide_tab(self):
        """Construit l'onglet Aide : guide d'utilisation scrollable avec sections structurées."""
        outer = ttk.Frame(self.tab_aide)
        outer.pack(fill="both", expand=True)

        sb = ttk.Scrollbar(outer, orient="vertical")
        sb.pack(side="right", fill="y")

        canvas = tk.Canvas(outer, bg="#f7f9fc", highlightthickness=0,
                           yscrollcommand=sb.set)
        canvas.pack(fill="both", expand=True)
        sb.config(command=canvas.yview)

        inner = ttk.Frame(canvas)
        win = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))

        def section(title, body):
            tk.Frame(inner, height=1, bg="#d0d8e8").pack(fill="x", padx=40, pady=(20, 6))
            tk.Label(inner, text=title, font=("Arial", 12, "bold"),
                     bg="#f7f9fc", fg="#2c3e50", anchor="w").pack(fill="x", padx=40)
            tk.Label(inner, text=body, font=("Arial", 10),
                     bg="#f7f9fc", fg="#333333", anchor="w", justify="left",
                     wraplength=900).pack(fill="x", padx=40, pady=(4, 0))

        tk.Label(inner, text="Guide d'utilisation — Rfill", font=("Arial", 16, "bold"),
                 bg="#f7f9fc", fg="#1a2a4a").pack(anchor="w", padx=40, pady=(30, 4))
        tk.Label(inner,
                 text="Importez des fichiers AutoCAD ou GeoGex, catégorisez les surfaces, puis exportez un tableau Excel normé.",
                 font=("Arial", 10), bg="#f7f9fc", fg="#666666").pack(anchor="w", padx=40)

        section("1 · Importer des fichiers",
                "Cliquez sur « AutoCAD (.xls) » ou « GeoGex (.xlsx) » pour sélectionner un ou plusieurs "
                "fichiers. Les données de chaque fichier s'accumulent dans la session. Le journal affiche "
                "le résultat de chaque import avec horodatage.\n\n"
                "• Bouton « Effacer tout » : supprime toutes les données de la session.\n"
                "• Les lignes de données sont détectées automatiquement grâce au numéro de séquence "
                "AutoCAD en colonne 0.\n"
                "• Le type de surface (SUB, SU, SHO…) est lu dans la colonne Calque.")

        section("2 · Glossaire",
                "Chaque type de surface détecté (SUB, SU, GLA…) possède son propre éditeur de catégories.\n\n"
                "• Sélectionnez un type dans la liste déroulante.\n"
                "• Les affectations sont automatiquement classées selon le glossaire interne.\n"
                "• Glissez-déposez une affectation (étiquette blanche) vers une autre catégorie.\n"
                "• Double-clic sur un titre de catégorie pour le renommer.\n"
                "• Clic droit sur un titre pour déplacer la catégorie vers une autre sur-catégorie.\n"
                "• ✕ supprime une catégorie (ses affectations passent dans « Non classé »).\n"
                "• « Sauvegarder » enregistre les nouvelles affectations dans le glossaire interne "
                "pour les prochaines sessions.")

        section("3 · Générer",
                "Renseignez les informations du projet (bâtiment, adresse, propriétaire, cadastre, mesurage…) "
                "puis cliquez sur « Générer Excel + HTML ».\n\n"
                "Deux fichiers sont créés côte à côte :\n"
                "• Un classeur Excel (.xlsx) — un onglet par type de surface, avec en-tête projet, "
                "tableau pivotant par étage/occupant/catégorie, sous-totaux et total général.\n"
                "• Un rapport HTML imprimable (.html) au format A4 paysage — un onglet par type, "
                "logos et tampon intégrés, prêt à imprimer via Ctrl+P.\n\n"
                "Le champ « Mesurage » remplit automatiquement la ligne « Nota : les superficies ont été "
                "calculées après mesurage des locaux en … » en bas du rapport HTML.")

        section("Formats de fichiers supportés",
                "• AutoCAD XLS : export AutoCAD natif (.xls). Les lignes de données ont un entier "
                "en colonne 0 (numéro de séquence). Le type de surface est dans la colonne Calque "
                "(ex. « SUB Contours » → type « SUB »).\n"
                "• GeoGex XLSX : export GeoGex (.xlsx), même structure attendue.")

        section("Codes de surface",
                "SUB  Surfaces Utiles Brutes\n"
                "SU   Surfaces Utiles\n"
                "SUBL Surfaces Utiles Brutes Locatives\n"
                "SUN  Surfaces Utiles Nettes\n"
                "SHO  Surfaces Hors Œuvre\n"
                "SDP  Surfaces De Plancher\n"
                "GLA  Surfaces Globales\n"
                "TAX  Synthèse des Surfaces Réelles (Art. 324 M ou Z Annexe III CGI)\n"
                "TSB  Tableau Surfaces Brutes")

        tk.Frame(inner, height=30, bg="#f7f9fc").pack()

    # =============================================================== À PROPOS

    def build_apropos_tab(self):
        """Construit l'onglet À propos : carte centrée avec version, auteur et copyright."""
        BG = "#f7f9fc"

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

        tk.Label(center, text=APP_NAME,
                 font=("Arial", 42, "bold"), fg="#1a2a4a", bg=BG).pack()

        tk.Frame(center, height=2, bg="#b0c4de", width=300).pack(pady=(6, 10))

        tk.Label(center, text="Analyse et génération de tableaux de surfaces",
                 font=("Arial", 12), fg="#555", bg=BG).pack()

        tk.Frame(center, height=20, bg=BG).pack()

        info_frame = tk.Frame(center, bg="#eef2f7", bd=1, relief="solid",
                              padx=30, pady=20)
        info_frame.pack()

        def row(label, value):
            r = tk.Frame(info_frame, bg="#eef2f7")
            r.pack(fill="x", pady=2)
            tk.Label(r, text=label, font=("Arial", 10, "bold"),
                     fg="#444", bg="#eef2f7", width=20, anchor="e").pack(side="left")
            tk.Label(r, text=value, font=("Arial", 10),
                     fg="#222", bg="#eef2f7", anchor="w").pack(side="left", padx=(8, 0))

        row("Version :", APP_VERSION)
        row("Build :", APP_BUILD)
        row("Mise à jour :", APP_DATE)
        row("Développé par :", APP_AUTHOR)

        tk.Frame(info_frame, height=12, bg="#eef2f7").pack()
        tk.Label(info_frame, text=APP_DESC, font=("Arial", 9), fg="#666",
                 bg="#eef2f7", justify="center").pack()

        tk.Frame(center, height=20, bg=BG).pack()

        tk.Label(center, text=APP_COPY,
                 font=("Arial", 9), fg="#999", bg=BG).pack()


if __name__ == "__main__":
    app = SurfaceApp()
    app.mainloop()
