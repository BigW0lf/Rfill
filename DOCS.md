# Documentation des fonctions — Rfill

## glossary_surf.py — Données de référence statiques

Module de données pures. Toutes les valeurs sont des constantes de référence chargées à l'import.

| Variable | Type | Description |
|---|---|---|
| `glossary_surf` | `dict[str, dict[str, list[str]]]` | Glossaires de mots-clés par famille de surface (`surfaces_utilisables`, `sho`, `spd`). Structure : `{glossaire: {catégorie: [keywords]}}`. |
| `denom_surf` | `dict[str, str]` | Code surface → clé de glossaire. Ex. `'SUB' → 'surfaces_utilisables'`. |
| `real_su_name` | `dict[str, str]` | Code surface → libellé complet pour les en-têtes. Ex. `'SU' → 'SURFACES UTILES (SU*)'`. |
| `superficie_names` | `dict[str, list[str]]` | Code surface → noms des super-catégories par défaut. |
| `predefined_cats` | `dict[str, list[str]]` | Code surface → catégories pré-saisies (SHO, SUBL, SUN, TAX, SDP). |
| `nota_surf` | `dict[str, list[str]]` | Code surface → notes réglementaires affichées en bas des tableaux. |

---

## process_surf.py — Traitement des données

### Fonctions privées

#### `_glossary_path() → str`
Retourne le chemin absolu de `glossary_surf.py`. En mode exe PyInstaller, pointe vers le dossier à côté de l'exécutable ; sinon vers le dossier du script.

#### `_reload_glossary()`
Recharge le module `glossary_surf` depuis le disque. En mode frozen, utilise `importlib.util` pour charger depuis le chemin absolu.

#### `_img_dir() → str`
Retourne le chemin du dossier `img/` adapté au mode d'exécution (PyInstaller `_MEIPASS` ou dossier du script).

#### `_img_to_base64(filename: str) → str`
Encode l'image `filename` (depuis `img/`) en data URI base64. Retourne `''` si le fichier est absent.

#### `_fmt_num(v) → str`
Formate une valeur numérique pour l'HTML : séparateur milliers = espace, décimale = virgule. Retourne `""` si `v` est `None` ou vide.

---

### Fonctions métier

#### `extract_info(file_path: str) → DataFrame`
Lit un fichier Excel AutoCAD/GeoGex et retourne un DataFrame normalisé.

Colonnes produites : `num_piece`, `Descriptif`, `Calque`, `N`, `Affectation`, `Occupant`, `Aire`, `Etage`, `Chambre`, `Lot`, `Accessibilité aux publics`, `type_su`.

- Filtre les lignes dont la colonne 0 est numérique (numéro de séquence AutoCAD).
- `type_su` = premier mot du champ `Calque` (ex. `"SUB Contours"` → `"SUB"`).

---

#### `tab_cd_type(df: DataFrame) → list[DataFrame]`
Agrège le DataFrame par type de surface, puis par `(Etage, Affectation, Occupant)` (somme des `Aire`).

**Retourne** : une liste de DataFrames, un par valeur unique de `type_su`.

---

#### `build_affectation_mapping(df_t: DataFrame, type_su: str) → DataFrame`
Construit un mapping `Affectation → catégorie` par lookup exact dans le glossaire.

Les affectations absentes du glossaire reçoivent la catégorie `"autres"`.

**Retourne** : `DataFrame` avec colonnes `['Affectation', 'cat']`.

---

#### `TCD2Tab(df_t, type_su, mapping_df=None) → tuple[DataFrame, DataFrame]`
Fusionne le DataFrame de surfaces avec le mapping de catégories et agrège par `(Etage, Occupant, cat)`.

- `mapping_df` : si `None`, calculé via `build_affectation_mapping`.

**Retourne** : `(df_tcd, mapping_df)`.

---

#### `Tab_output(df_tcd, infos, super_cat_map=None) → dict`
Transforme `df_tcd` en tableaux finaux avec pivot, sous-totaux par étage et total général.

- Crée un pivot `(Etage[, Occupant]) × catégorie`.
- Trie les occupants vides en dernier au sein de chaque étage.
- Insère des lignes de sous-total par étage si plusieurs occupants.
- Ajoute une ligne `TOTAL` en bas.
- Si `super_cat_map` fourni : ajoute des colonnes `[sc_name]` de sous-total et calcule `sc_spans` pour l'export Excel.

**Paramètres**
- `infos` : `dict` avec les clés `batiment`, `adresse`, `proprio`, `cadastre`, `date`, `dossier`, `mesurage`.
- `super_cat_map` : `{type_su: {sc_name: [categories]}}`.

**Retourne** : `{type_su: {"info": [str], "data": DataFrame, "sc_spans": dict, "nota": [str]}}`.

---

#### `export_tables_to_excel(output_tables, output_path)`
Écrit le classeur Excel multi-feuilles (une feuille par type de surface).

Mise en forme : en-tête projet en italique, en-têtes sur-catégories en bleu acier (`#B0C4DE`), en-têtes colonnes en bleu clair (`#DDEEFF`), totaux par étage en `#DCE6F1`, grand total en `#B8CCE4`, colonnes sous-total en `#E8F0E8`. Notes réglementaires en bas, fusionnées sur toute la largeur.

---

#### `update_glossary(mapping_df, type_su) → bool`
Ajoute dans `glossary_surf.py` les nouvelles affectations classifiées (hors `"autres"`).

- Valide la syntaxe via `ast.parse` avant écriture.
- Écriture atomique : fichier temporaire → `shutil.move`.
- Recharge le module après modification via `_reload_glossary`.

**Retourne** : `True` si le fichier a été modifié.

---

#### `export_tables_to_html(output_tables, infos, html_path)`
Génère un rapport HTML auto-contenu, imprimable A4 paysage (une `<section>` par type de surface).

- Logos et tampon encodés en base64 (aucune ressource externe).
- En-tête : logos gauche · infos projet centre · date/dossier droite.
- Navigation par onglets (écran uniquement, `@media print` masque la `<nav>`).
- Pied de page : nota intro + notes réglementaires + tampon.

---

### Fonctions internes de rendu HTML (dans `export_tables_to_html`)

#### `render_header(sheet_title) → str`
Génère le bloc `<header>` HTML avec logos, infos projet centrées et métadonnées à droite.

#### `render_table(df, sc_spans) → str`
Génère le `<table class="surf">` avec colgroup, en-têtes sur-catégories, en-têtes colonnes et lignes de données (classes CSS `total-etage`, `total`).

---

## Main.py — Interface graphique (Tkinter)

### Classe `SurfaceApp(tk.Tk)`

Application principale Rfill. Fenêtre 1200×780, 5 onglets : Importer · Glossaire · Générer · Aide · À propos.

#### État interne

| Attribut | Type | Description |
|---|---|---|
| `all_df` | `DataFrame \| None` | Données brutes cumulées (toutes sources importées) |
| `df_types` | `list[DataFrame] \| None` | Liste agrégée par type de surface (`tab_cd_type`) |
| `mappings` | `dict[str, DataFrame]` | `{type_su: mapping_df}` |
| `empty_categories` | `dict[str, set[str]]` | Catégories vides créées manuellement |
| `super_cats` | `dict[str, dict[str, list[str]]]` | `{type_su: {sc_name: [cats]}}` |
| `saved_glossaries` | `set[str]` | Types sauvegardés dans le glossaire cette session |
| `columns` | `dict[str, tk.Frame]` | Widgets de colonnes de catégories dans le board |

---

### Méthodes de configuration

#### `__init__()`
Initialise la fenêtre, les styles ttk, les 5 onglets et toutes les variables d'état.

#### `_setup_styles()`
Configure le thème `clam` et les styles ttk (polices Arial, couleurs).

---

### Méthodes utilitaires

#### `_ts() → str`
Retourne `HH:MM:SS` pour les entrées de journal.

#### `_log(widget, msg, tag="info")`
Insère `[HH:MM:SS]  msg` dans le `Text` widget avec le tag de couleur (`ok` vert, `warn` jaune, `err` rouge, `info` bleu, `sep` gris).

#### `_log_sep(widget, label="")`
Insère un séparateur visuel dans le journal (60 tirets, avec label centré si fourni).

---

### Onglet 1 — Importer

#### `build_import_tab()`
Construit l'onglet : boutons *AutoCAD (.xls)* / *GeoGex (.xlsx)* / *Effacer tout*, console sombre (Consolas 9).

#### `load_files(source: str)`
Importe les fichiers sélectionnés (`"autocad"` ou `"geogex"`). Appelle `extract_info` puis `tab_cd_type`, accumule dans `all_df`, déclenche `update_glossaire_tab`.

#### `clear_data()`
Efface toutes les données après confirmation : remet à zéro `all_df`, `df_types`, `mappings`, `empty_categories`, le board et la combobox.

---

### Onglet 2 — Glossaire

#### `build_glossaire_tab()`
Construit l'onglet : combobox de type, boutons *+ Ajouter* / *Sauvegarder*, canvas scrollable (H+V) contenant le `board`.

#### `_on_canvas_resize(event)`
Adapte la largeur du frame `board` à la largeur courante du canvas.

#### `_on_tab_changed(event=None)`
Affiche *"Importez d'abord des fichiers"* dans le board si aucun mapping n'est chargé.

#### `update_glossaire_tab()`
Recalcule `mappings` et `super_cats` depuis `df_types`, met à jour la combobox. Appelée automatiquement après chaque import.

#### `load_glossaire_board(event=None)`
Reconstruit entièrement le board des catégories pour le type sélectionné : frames de sur-catégories, colonnes de catégories, widgets d'affectation.

#### `_make_aff_widget(parent, aff, cat)`
Crée un label blanc draggable représentant une affectation dans sa colonne de catégorie.

---

### Drag & drop — Affectations (labels blancs)

#### `start_drag(event, aff, cat)`
Démarre le drag : mémorise `drag_item` / `drag_origin_cat`, crée le label fantôme orange.

#### `do_drag(event)`
Met à jour la position du fantôme en suivant le curseur.

#### `stop_drag(event)`
Dépose l'affectation dans la catégorie cible (hit-test sur `self.columns`), met à jour `mapping_df`, rafraîchit le board.

---

### Drag & drop — Colonnes de catégories (réordonnancement)

#### `_cat_drag_start(event, cat, sc)`
Initialise le contexte de drag avec la catégorie source et les coordonnées de départ.

#### `_cat_drag_do(event)`
Active le mode drag après 6 px, crée et déplace le fantôme orange foncé.

#### `_cat_drag_stop(event)`
Insère la catégorie source avant la cible dans `super_cats`, rafraîchit le board.

---

### Gestion des catégories

#### `show_supercat_menu(event, cat, current_sc)`
Menu contextuel listant les sur-catégories disponibles pour déplacer `cat`.

#### `move_cat_to_supercat(cat, from_sc, to_sc)`
Retire `cat` de `from_sc`, l'ajoute à `to_sc`, rafraîchit le board.

#### `rename_category(old_name)`
Renomme une catégorie via dialog, met à jour `mappings`, `empty_categories` et `super_cats`.

#### `add_category()`
Crée une catégorie vide via dialog, l'ajoute à `empty_categories` et à la première sur-catégorie.

#### `delete_category(cat)`
Supprime une catégorie après confirmation, reporte ses affectations vers `"autres"`.

#### `save_glossaire()`
Appelle `update_glossary` pour persister les affectations classifiées dans `glossary_surf.py`.

---

### Gestion des sur-catégories

#### `rename_supercat(old_sc, type_su)`
Renomme une sur-catégorie en reconstruisant `super_cats[type_su]` avec la nouvelle clé à la même position.

---

### Onglet 3 — Générer

#### `build_generer_tab()`
Construit l'onglet : formulaire projet (7 champs), bouton *Générer Excel + HTML*, case *Ouvrir après*, console.

#### `generate_final_tables()`
Export complet :
1. Sauvegarde auto des glossaires non sauvegardés.
2. Sélection du fichier de destination `.xlsx`.
3. Pour chaque type : `TCD2Tab` → `Tab_output`.
4. `export_tables_to_excel` + `export_tables_to_html`.
5. Ouverture automatique du fichier Excel si la case est cochée.

---

### Onglets informatifs

#### `build_aide_tab()`
Guide d'utilisation scrollable avec sections structurées (import, glossaire, génération, codes de surface).

#### `build_apropos_tab()`
Carte centrée sur un canvas avec nom de l'application, version, auteur et copyright.
