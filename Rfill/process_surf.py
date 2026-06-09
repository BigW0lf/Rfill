import ast
import base64
import html
import os
import pprint
import shutil
import sys
import tempfile
import importlib
import importlib.util

import pandas as pd
import glossary_surf

# En mode exe PyInstaller, si l'utilisateur a un glossary_surf.py modifié à côté
# de l'exe, le charger immédiatement en priorité sur la version bundlée.
def _try_load_user_glossary():
    if not getattr(sys, "frozen", False):
        return
    user_path = os.path.join(os.path.dirname(sys.executable), "glossary_surf.py")
    if not os.path.isfile(user_path):
        return
    global glossary_surf
    spec = importlib.util.spec_from_file_location("glossary_surf", user_path)
    mod  = importlib.util.module_from_spec(spec)
    sys.modules["glossary_surf"] = mod
    spec.loader.exec_module(mod)
    glossary_surf = mod

_try_load_user_glossary()


# ─── Correspondance numéro d'étage → libellé affiché dans les tableaux ────────
_ETAGE_LABELS = {
    -3.0: "3ème Sous-sol",
    -2.0: "2ème Sous-sol",
    -1.0: "1er Sous-sol",
    -0.5: "Rez-de-jardin",
     0.0: "Rez-de-chaussée",
     0.5: "1er Entresol",
     1.0: "1er Étage",
     1.5: "2ème Entresol",
     2.0: "2ème Étage",
     2.5: "3ème Entresol",
     3.0: "3ème Étage",
     3.5: "4ème Entresol",
     4.0: "4ème Étage",
     4.5: "5ème Entresol",
     5.0: "5ème Étage",
     5.5: "6ème Entresol",
     6.0: "6ème Étage",
     6.5: "7ème Entresol",
     7.0: "7ème Étage",
     7.5: "8ème Entresol",
     8.0: "8ème Étage",
}

def _etage_label(val):
    """Convertit une valeur numérique d'étage en libellé français (ex. -1 → '1er Sous-sol')."""
    if str(val) == "TOTAL":
        return val
    try:
        f = float(val)
    except (TypeError, ValueError):
        return str(val)
    if f in _ETAGE_LABELS:
        return _ETAGE_LABELS[f]
    n = int(f)
    if f == n:
        if n < 0:
            abs_n = abs(n)
            suffix = "er" if abs_n == 1 else "ème"
            return f"{abs_n}{suffix} Sous-sol"
        elif n == 0:
            return "Rez-de-chaussée"
        else:
            suffix = "er" if n == 1 else "ème"
            return f"{n}{suffix} Étage"
    else:
        entresol_n = int(f + 0.5)
        suffix = "er" if entresol_n == 1 else "ème"
        return f"{entresol_n}{suffix} Entresol"


def _glossary_path():
    """Retourne le chemin absolu de ``glossary_surf.py``.

    En mode exe PyInstaller, le fichier est situé à côté de l'exécutable
    (modifiable par l'utilisateur) ; en mode script, à côté de ce module.

    Returns:
        str: Chemin absolu vers ``glossary_surf.py``.
    """
    if getattr(sys, 'frozen', False):
        # En mode exe : dossier à côté de Rfill.exe
        return os.path.join(os.path.dirname(sys.executable), "glossary_surf.py")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "glossary_surf.py")


def _reload_glossary():
    """Recharge le module ``glossary_surf`` depuis le fichier sur disque.

    En mode *frozen* (PyInstaller), utilise :func:`importlib.util.spec_from_file_location`
    pour charger depuis le chemin absolu renvoyé par :func:`_glossary_path`.
    En mode script, délègue à :func:`importlib.reload`.
    """
    global glossary_surf
    if getattr(sys, 'frozen', False):
        path = _glossary_path()
        spec = importlib.util.spec_from_file_location("glossary_surf", path)
        mod  = importlib.util.module_from_spec(spec)
        sys.modules["glossary_surf"] = mod
        spec.loader.exec_module(mod)
        glossary_surf = mod
    else:
        importlib.reload(glossary_surf)

# =========================================================
#  LOGIQUE METIER
# =========================================================

def extract_info(file_path):
    """Lit un fichier Excel AutoCAD/GeoGex et retourne un DataFrame normalisé.

    Filtre les lignes dont la colonne 0 est un entier (numéro de séquence AutoCAD)
    et dérive ``type_su`` depuis le premier mot de la colonne ``Calque``
    (ex. ``"SUB Contours"`` → ``"SUB"``).

    Args:
        file_path (str): Chemin vers le fichier ``.xls`` ou ``.xlsx``.

    Returns:
        pandas.DataFrame: Colonnes : ``num_piece``, ``Descriptif``, ``Calque``,
        ``N``, ``Affectation``, ``Occupant``, ``Aire``, ``Etage``, ``Chambre``,
        ``Lot``, ``Accessibilité aux publics``, ``type_su``.

    Raises:
        ValueError: Si le fichier ne contient aucune ligne de données valides.
    """
    _EXPECTED_COLS = ["num_piece", "Descriptif", "Calque", "N", "Affectation",
                      "Occupant", "Aire", "Etage", "Chambre", "Lot",
                      "Accessibilité aux publics"]

    # GeoGex peut exporter avec une ligne d'en-tête — on cherche la ligne contenant "Calque"
    df_raw = pd.read_excel(file_path, header=None)

    header_row = None
    for i, row in df_raw.iterrows():
        if any(str(v).strip().lower() == "calque" for v in row):
            header_row = i
            break

    if header_row is not None:
        # Relire avec l'en-tête détecté
        df = pd.read_excel(file_path, header=header_row)
        # Renommer les colonnes en noms normalisés par position
        if len(df.columns) >= len(_EXPECTED_COLS):
            rename_map = {df.columns[i]: _EXPECTED_COLS[i] for i in range(len(_EXPECTED_COLS))}
            df = df.rename(columns=rename_map)
        # Conserver uniquement les colonnes attendues
        df = df[[c for c in _EXPECTED_COLS if c in df.columns]].copy()
    else:
        df = df_raw.copy()

    # Les lignes de données ont un entier en colonne 0 (numéro de séquence AutoCAD)
    col0 = df.columns[0]
    df[col0] = pd.to_numeric(df[col0], errors="coerce")
    df = df[df[col0].notna()].copy().reset_index(drop=True)

    if df.empty:
        raise ValueError("Aucune ligne de données valides trouvée dans le fichier.")

    # S'assurer qu'on a exactement les 11 colonnes attendues (compléter si manquant)
    if list(df.columns) != _EXPECTED_COLS:
        # Essayer d'assigner par position
        n = min(len(df.columns), len(_EXPECTED_COLS))
        col_map = {df.columns[i]: _EXPECTED_COLS[i] for i in range(n)}
        df = df.rename(columns=col_map)
        for c in _EXPECTED_COLS:
            if c not in df.columns:
                df[c] = ""
        df = df[_EXPECTED_COLS]

    # Type de surface directement depuis le calque : "SUB Contours" → "SUB"
    df["type_su"] = df["Calque"].astype(str).str.strip().str.split().str[0].str.upper()

    return df


def tab_cd_type(df):
    """Agrège le DataFrame brut par type de surface, puis par *(Etage, Affectation, Occupant)*.

    Pour chaque valeur unique de ``type_su``, les surfaces sont sommées par groupe.
    Les valeurs manquantes de ``Occupant`` sont remplacées par ``""`` ;
    ``Aire`` est converti en numérique (valeurs non parsables → 0).

    Args:
        df (pandas.DataFrame): DataFrame produit par :func:`extract_info`.

    Returns:
        list[pandas.DataFrame]: Un DataFrame par type de surface, chacun avec les
        colonnes ``Etage``, ``Affectation``, ``Occupant``, ``Aire``, ``type_su``.
    """
    df_type = []
    type_unique = df.type_su.dropna().unique()

    for t in type_unique:
        df_t = df[df.type_su == t].copy()
        df_t["Occupant"] = df_t["Occupant"].fillna("").astype(str)
        df_t["Aire"] = pd.to_numeric(df_t["Aire"], errors="coerce").fillna(0)
        df_t = df_t.groupby(["Etage", "Affectation", "Occupant"]).agg({"Aire": "sum"}).reset_index()
        df_t["type_su"] = t
        df_type.append(df_t)

    return df_type


def build_affectation_mapping(df_t, type_su):
    """Construit un mapping ``Affectation → catégorie`` par lookup exact dans le glossaire.

    La résolution passe par :data:`glossary_surf.denom_surf` pour trouver le
    glossaire applicable, puis effectue une comparaison normalisée (minuscules,
    espaces supprimés) entre chaque affectation et les mots-clés du glossaire.
    Les affectations sans correspondance reçoivent la catégorie ``"autres"``.

    Args:
        df_t (pandas.DataFrame): Données pour un type de surface unique,
            avec au moins la colonne ``Affectation``.
        type_su (str): Code de surface (ex. ``"SUB"``, ``"SHO"``).

    Returns:
        pandas.DataFrame: Colonnes ``['Affectation', 'cat']``,
        une ligne par affectation unique présente dans *df_t*.
    """
    type_su_glo = glossary_surf.denom_surf.get(type_su)
    if not type_su_glo or type_su_glo not in glossary_surf.glossary_surf:
        df_aff = df_t[["Affectation"]].drop_duplicates().copy()
        df_aff["cat"] = "autres"
        return df_aff
    type_glossary = glossary_surf.glossary_surf[type_su_glo]

    # Table de lookup exacte : keyword normalisé → catégorie
    keyword_to_cat = {
        k.lower().strip(): cat
        for cat, keywords in type_glossary.items()
        for k in keywords
    }

    df_aff = df_t[["Affectation"]].drop_duplicates().copy()
    df_aff["cat"] = df_aff["Affectation"].apply(
        lambda x: keyword_to_cat.get(str(x).lower().strip(), "autres")
    )
    return df_aff


def TCD2Tab(df_t, type_su, mapping_df=None):
    """Fusionne les surfaces avec leur mapping de catégories et agrège par *(Etage, Occupant, cat)*.

    Args:
        df_t (pandas.DataFrame): DataFrame d'un type de surface unique,
            colonnes ``['Etage', 'Affectation', 'Occupant', 'Aire', 'type_su']``.
        type_su (str): Code de surface (ex. ``"SU"``).
        mapping_df (pandas.DataFrame, optional): Table ``['Affectation', 'cat']``
            issue de l'interface utilisateur. Si ``None``, calculée
            automatiquement via :func:`build_affectation_mapping`.

    Returns:
        tuple[pandas.DataFrame, pandas.DataFrame]:
            - **df_tcd** — pivot agrégé, colonnes
              ``['Etage', 'Occupant', 'cat', 'Aire', 'type_su']``.
            - **mapping_df** — mapping effectivement utilisé (utile si calculé automatiquement).
    """
    if mapping_df is None:
        mapping_df = build_affectation_mapping(df_t, type_su)

    df = df_t.merge(mapping_df, on="Affectation", how="left")
    df["cat"] = df["cat"].fillna("autres")

    df_tcd = (
        df.groupby(["Etage", "Occupant", "cat"])
          .agg({"Aire": "sum"})
          .reset_index()
    )
    df_tcd["type_su"] = type_su

    return df_tcd, mapping_df



# Colonnes de données SDP dans l'ordre normalisé
_INFO_LABELS = {
    "batiment": "Bâtiment",
    "adresse":  "Adresse",
    "proprio":  "Propriétaire",
    "cadastre": "Cadastre",
    "date":     "Date",
    "dossier":  "Dossier",
    "mesurage": "Mesurage",
}


def _build_info_lines(type_su, infos):
    """Construit les lignes d'en-tête projet en omettant les champs vides."""
    lines = [glossary_surf.real_su_name.get(type_su, type_su)]
    for key, label in _INFO_LABELS.items():
        val = infos.get(key, "").strip()
        if val:
            lines.append(f"{label} : {val}")
    lines.append("")
    return lines


_SDP_DATA_COLS = [
    "Planchers avant déductions",
    "Vides Gaines Techniques",
    "Surfaces avec h < 1.80 m",
    "Stationnements",
    "Combles non aménageables",
    "Locaux techniques",
    "Caves/annexes",
    "Déduction 10% (Habitation)",
]

# Numéro de référence affiché au-dessus de chaque colonne (correspond aux nota)
_SDP_COL_NUMBERS = {
    "Planchers avant déductions":  "(1)",
    "Vides Gaines Techniques":     "(2)",
    "Surfaces avec h < 1.80 m":    "(3)",
    "Total TA":                    "(4)",
    "Stationnements":              "(5)",
    "Combles non aménageables":    "(6)",
    "Locaux techniques":           "(7)",
    "Caves/annexes":               "(8)",
    "Déduction 10% (Habitation)":  "(9)",
    "Total":                       "(10)",
}


def _build_sdp_table(df_t, infos):
    """Construit le tableau SDP figé avec colonnes calculées Total TA et SDP.

    Structure :
        Etage | (1) | (2) | (3) | Total TA | (5) | (6) | (7) | (8) | (9) | SDP

    Total TA  = (1) - (2) - (3)
    SDP       = Total TA - (5) - (6) - (7) - (8) - (9)
    """
    has_occupant = (
        "Occupant" in df_t.columns
        and df_t["Occupant"].str.strip().ne("").any()
    )
    index_cols = ["Etage", "Occupant"] if has_occupant else ["Etage"]

    table = df_t.pivot_table(
        index=index_cols,
        columns="cat",
        values="Aire",
        aggfunc="sum",
        fill_value=0,
    )
    table.columns.name = None
    table = table.reset_index()

    # Ajouter les colonnes manquantes à 0
    for col in _SDP_DATA_COLS:
        if col not in table.columns:
            table[col] = 0.0

    deductions = ["Vides Gaines Techniques", "Surfaces avec h < 1.80 m"]
    table["Total TA"] = (
        table["Planchers avant déductions"]
        - table[deductions].sum(axis=1)
    )

    reste_cols = ["Stationnements", "Combles non aménageables",
                  "Locaux techniques", "Caves/annexes", "Déduction 10% (Habitation)"]
    table["Total"] = table["Total TA"] - table[reste_cols].sum(axis=1)

    ordered = index_cols + [
        "Planchers avant déductions",
        "Vides Gaines Techniques",
        "Surfaces avec h < 1.80 m",
        "Total TA",
        "Stationnements",
        "Combles non aménageables",
        "Locaux techniques",
        "Caves/annexes",
        "Déduction 10% (Habitation)",
        "Total",
    ]
    table = table[[c for c in ordered if c in table.columns]]

    all_value_cols = [c for c in table.columns if c not in index_cols]

    # Sous-totaux par étage si multi-occupants
    if has_occupant:
        etages = list(dict.fromkeys(table["Etage"]))
        chunks = []
        for etage in etages:
            chunk = table[table["Etage"] == etage].copy()
            chunks.append(chunk)
            sub = {c: "" for c in table.columns}
            sub["Etage"] = str(etage)
            sub["Occupant"] = "— Total étage"
            for c in all_value_cols:
                sub[c] = pd.to_numeric(chunk[c], errors="coerce").sum()
            chunks.append(pd.DataFrame([sub]))
        table = pd.concat(chunks, ignore_index=True)

    total_row = {c: "" for c in table.columns}
    total_row["Etage"] = "TOTAL"
    mask_not_subtotal = (
        table.get("Occupant", pd.Series([""] * len(table))) != "— Total étage"
    )
    for c in all_value_cols:
        total_row[c] = pd.to_numeric(table.loc[mask_not_subtotal, c], errors="coerce").sum()
    table = pd.concat([table, pd.DataFrame([total_row])], ignore_index=True)

    # Convertir les valeurs d'étage en libellés français
    table["Etage"] = table["Etage"].apply(
        lambda v: v if str(v) in ("TOTAL", "— Total étage") else _etage_label(v)
    )

    nota = glossary_surf.nota_surf.get("SDP", [])
    return {"info": _build_info_lines("SDP", infos), "data": table, "sc_spans": {},
            "nota": nota, "sdp_fixed": True, "col_numbers": _SDP_COL_NUMBERS,
            "mesurage": infos.get("mesurage", "").strip()}


def Tab_output(df_tcd, infos, super_cat_map=None):
    """Transforme *df_tcd* en tableaux finaux prêts à l'export.

    Pour chaque type de surface, construit un tableau croisé
    *(Etage [× Occupant]) × catégorie* avec :

    * colonnes de sous-total par sur-catégorie si *super_cat_map* est fourni ;
    * lignes de sous-total par étage (multi-occupants uniquement) ;
    * ligne ``TOTAL`` globale en bas.

    Args:
        df_tcd (pandas.DataFrame): Sortie de :func:`TCD2Tab`, colonnes
            ``['Etage', 'Occupant', 'cat', 'Aire', 'type_su']``.
        infos (dict): Métadonnées projet — clés attendues : ``batiment``,
            ``adresse``, ``proprio``, ``cadastre``, ``date``, ``dossier``,
            ``mesurage``.
        super_cat_map (dict, optional): ``{type_su: {sc_name: [categories]}}``.
            Définit l'ordre des colonnes et les sous-totaux par sur-catégorie.

    Returns:
        dict: ``{type_su: {"info": list[str], "data": DataFrame,
        "sc_spans": dict, "nota": list[str]}}``.

        * ``info`` — lignes d'en-tête projet à écrire avant le tableau.
        * ``data`` — tableau final incluant les lignes de total.
        * ``sc_spans`` — ``{sc_name: (col_start, col_end)}`` pour la mise en
          forme Excel des sur-catégories.
        * ``nota`` — notes réglementaires issues de :data:`glossary_surf.nota_surf`.
    """
    output_tables = {}
    type_su_list = pd.unique(df_tcd["type_su"].values)

    # Catégories prédéfinies dans le glossaire pour chaque type (ordre de référence)
    from glossary_surf import predefined_cats as _predefined_cats

    for t in type_su_list:
        df_t = df_tcd[df_tcd["type_su"] == t].copy()
        df_t = df_t[df_t["cat"] != "autres"]

        if t == "SDP":
            output_tables["SDP"] = _build_sdp_table(df_t, infos)
            continue

        has_occupant = (
            "Occupant" in df_t.columns
            and df_t["Occupant"].str.strip().ne("").any()
        )
        index_cols = ["Etage", "Occupant"] if has_occupant else ["Etage"]

        table = df_t.pivot_table(
            index=index_cols,
            columns="cat",
            values="Aire",
            aggfunc="sum",
            fill_value=0
        )
        table.columns.name = None
        table = table.reset_index()

        # Occupants vides triés en dernier au sein de chaque étage
        if has_occupant and "Occupant" in table.columns:
            occ_empty = (table["Occupant"].fillna("").str.strip() == "").astype(int)
            table = (
                table.assign(_s=occ_empty)
                .sort_values(["Etage", "_s", "Occupant"])
                .drop(columns=["_s"])
                .reset_index(drop=True)
            )

        # Ajouter les catégories prédéfinies manquantes (affichées à 0 si vides)
        # SAUF si elles ont été supprimées (non présentes dans super_cat_map)
        existing_cats = set(table.columns) - set(index_cols)
        sc_all_cats = set()
        if super_cat_map and t in super_cat_map:
            for cats in super_cat_map[t].values():
                sc_all_cats.update(cats)
        predef = _predefined_cats.get(t, [])
        cats_to_add = [
            c for c in predef
            if c not in existing_cats
            and (not sc_all_cats or c in sc_all_cats)
        ]
        for c in cats_to_add:
            table[c] = 0.0

        cat_cols = [c for c in table.columns if c not in index_cols]
        table["Total"] = table[cat_cols].sum(axis=1)

        # Sous-totaux par sur-catégorie
        sc_map = (super_cat_map or {}).get(t, {})
        sc_spans = {}  # {sc_name: (col_start_idx, col_end_idx)} dans le df final

        # Pour SUB : "Terrasses /\nBalcons" est hors sous-total, placée après [Superficies annexes]
        _SUB_STANDALONE = "Terrasses /\nBalcons"
        standalone_cols = []
        if t == "SUB" and _SUB_STANDALONE in cat_cols:
            standalone_cols = [_SUB_STANDALONE]

        if sc_map:
            ordered = list(index_cols)
            for sc_name, sc_cats in sc_map.items():
                # Exclure les colonnes standalone du sous-total
                sc_actual = [c for c in sc_cats if c in cat_cols and c not in standalone_cols]
                sc_alone  = [c for c in standalone_cols if c in sc_cats and c in cat_cols]
                if sc_actual:
                    ordered.extend(sc_actual)
                    sub_col = f"[{sc_name}]"
                    table[sub_col] = table[sc_actual].sum(axis=1)
                    ordered.append(sub_col)
                # Colonnes standalone après le sous-total (ou seules si pas de sc_actual)
                ordered.extend(sc_alone)
            remaining = [c for c in cat_cols if c not in ordered and c not in standalone_cols]
            ordered.extend(remaining)
            ordered.append("Total")
            table = table[[c for c in ordered if c in table.columns]]
            # Calculer les spans de colonnes pour l'export Excel
            final_cols = list(table.columns)
            for sc_name, sc_cats in sc_map.items():
                sc_actual = [c for c in sc_cats if c in cat_cols and c not in standalone_cols]
                sc_standalone = [c for c in sc_cats if c in standalone_cols and c in cat_cols]
                if sc_actual:
                    sub_col = f"[{sc_name}]"
                    end_col = sc_standalone[0] if sc_standalone else sub_col
                    sc_spans[sc_name] = (final_cols.index(sc_actual[0]),
                                         final_cols.index(end_col))

        all_value_cols = [c for c in table.columns if c not in index_cols]

        # Sous-totaux par étage — toujours générés
        etages = list(dict.fromkeys(table["Etage"]))
        chunks = []
        for etage in etages:
            chunk = table[table["Etage"] == etage].copy()
            chunks.append(chunk)
            sub = {c: "" for c in table.columns}
            sub["Etage"] = f"_sub_{etage}"  # marqueur interne avec le nom d'étage
            if has_occupant:
                sub["Occupant"] = "— Total étage"
            for c in all_value_cols:
                sub[c] = pd.to_numeric(chunk[c], errors="coerce").sum()
            chunks.append(pd.DataFrame([sub]))
        table = pd.concat(chunks, ignore_index=True)

        total_row = {c: "" for c in table.columns}
        total_row["Etage"] = "TOTAL"
        not_subtotal = ~table["Etage"].astype(str).str.startswith("_sub_")
        for c in all_value_cols:
            total_row[c] = pd.to_numeric(table.loc[not_subtotal, c], errors="coerce").sum()
        table = pd.concat([table, pd.DataFrame([total_row])], ignore_index=True)

        # Convertir les valeurs d'étage en libellés français
        table["Etage"] = table["Etage"].apply(
            lambda v: v if str(v) == "TOTAL" else (
                _etage_label(str(v)[5:]) + " — Total étage" if str(v).startswith("_sub_") else _etage_label(v)
            )
        )

        nota = glossary_surf.nota_surf.get(t, [])
        output_tables[t] = {"info": _build_info_lines(t, infos), "data": table, "sc_spans": sc_spans,
                             "nota": nota, "mesurage": infos.get("mesurage", "").strip()}

    return output_tables


def verify_totals(df_types, output_tables, mappings=None):
    """Compare les totaux du fichier source avec les tableaux générés.

    Vérifie par étage et au global que :
    - entrée classée (hors "autres") == sortie tableau  → OK
    - sortie > entrée classée                          → anomalie réelle (warn)
    - entrée classée > sortie                          → anomalie réelle (warn)
    - surfaces non classées (autres)                   → info seulement

    Args:
        df_types (list[DataFrame]): Sortie de :func:`tab_cd_type`.
        output_tables (dict): Sortie de :func:`Tab_output`.
        mappings (dict, optional): ``{type_su: DataFrame['Affectation','cat']}``
            depuis l'interface. Permet de savoir exactement quelles affectations
            sont "autres". Si absent, toute la différence entrée-sortie est
            considérée comme "autres".

    Returns:
        list[dict]: clés ``type``, ``niveau``, ``entree``, ``entree_classee``,
        ``autres``, ``sortie``, ``ecart``, ``info`` (bool), ``msg``.
    """
    anomalies = []
    mappings = mappings or {}

    for df_t in df_types:
        t = df_t["type_su"].iloc[0]
        if t not in output_tables:
            continue

        # ── Identifier les affectations "autres" via le mapping ───────────────
        mapping_df = mappings.get(t)
        if mapping_df is not None:
            affectations_autres = set(
                mapping_df.loc[mapping_df["cat"] == "autres", "Affectation"].tolist()
            )
        else:
            affectations_autres = set()  # inconnu — on utilisera la différence

        # SDP a part : sa structure soustractive la rend non comparable entrée/sortie
        if output_tables[t].get("sdp_fixed"):
            continue

        tbl = output_tables[t]["data"]
        index_cols = [c for c in ("Etage", "Occupant") if c in tbl.columns]
        value_cols = [c for c in tbl.columns
                      if c not in index_cols
                      and not (str(c).startswith("[") and str(c).endswith("]"))
                      and c != "Total"]

        # ── Par étage ─────────────────────────────────────────────────────────
        for etage in df_t["Etage"].unique():
            df_etage = df_t[df_t["Etage"] == etage]
            sum_in       = float(pd.to_numeric(df_etage["Aire"], errors="coerce").sum())
            if affectations_autres:
                sum_autres   = float(pd.to_numeric(
                    df_etage.loc[df_etage["Affectation"].isin(affectations_autres), "Aire"],
                    errors="coerce").sum())
            else:
                sum_autres = 0.0
            sum_in_c = sum_in - sum_autres  # ce qui doit apparaître dans le tableau

            etage_label = _etage_label(etage)
            mask_tbl = (
                tbl["Etage"].astype(str) == str(etage_label)
            ) & (
                tbl.get("Occupant", pd.Series([""] * len(tbl))).astype(str) != "— Total étage"
            )
            rows_out = tbl[mask_tbl]
            sum_out  = float(rows_out[value_cols].apply(
                pd.to_numeric, errors="coerce").sum().sum()) if not rows_out.empty else 0.0

            ecart = sum_in_c - sum_out
            if abs(ecart) > 0.01:
                is_info = affectations_autres and abs(sum_in_c - sum_out) < 0.01
                anomalies.append({
                    "type": t, "niveau": str(etage_label),
                    "entree": round(sum_in, 2),
                    "entree_classee": round(sum_in_c, 2),
                    "autres": round(sum_autres, 2),
                    "sortie": round(sum_out, 2),
                    "ecart": round(ecart, 2),
                    "info": False,
                    "msg": (f"{t} | {etage_label} : "
                            f"classes={sum_in_c:.2f} m2 "
                            f"sortie={sum_out:.2f} m2 "
                            f"ecart={ecart:+.2f} m2"
                            + (f" (dont {sum_autres:.2f} m2 non classes)" if sum_autres > 0 else "")),
                })

        # ── Global ────────────────────────────────────────────────────────────
        total_in   = float(pd.to_numeric(df_t["Aire"], errors="coerce").sum())
        if affectations_autres:
            total_autres = float(pd.to_numeric(
                df_t.loc[df_t["Affectation"].isin(affectations_autres), "Aire"],
                errors="coerce").sum())
        else:
            total_autres = 0.0
        total_in_c = total_in - total_autres

        mask_data = (
            tbl["Etage"].astype(str) != "TOTAL"
        ) & (
            ~tbl["Etage"].astype(str).str.endswith("— Total étage")
        ) & (
            tbl.get("Occupant", pd.Series([""] * len(tbl))).astype(str) != "— Total étage"
        )
        total_out = float(tbl[mask_data][value_cols].apply(
            pd.to_numeric, errors="coerce").sum().sum())

        ecart = total_in_c - total_out

        if total_autres > 0.01 and abs(ecart) <= 0.01:
            # Tout est classé correctement, juste des "autres" exclus → info
            anomalies.append({
                "type": t, "niveau": "TOTAL",
                "entree": round(total_in, 2),
                "entree_classee": round(total_in_c, 2),
                "autres": round(total_autres, 2),
                "sortie": round(total_out, 2),
                "ecart": 0.0,
                "info": True,
                "msg": (f"{t} | TOTAL : {total_autres:.2f} m2 non classes "
                        f"exclus du tableau sur {total_in:.2f} m2 total"),
            })
        elif abs(ecart) > 0.01:
            anomalies.append({
                "type": t, "niveau": "TOTAL",
                "entree": round(total_in, 2),
                "entree_classee": round(total_in_c, 2),
                "autres": round(total_autres, 2),
                "sortie": round(total_out, 2),
                "ecart": round(ecart, 2),
                "info": False,
                "msg": (f"{t} | TOTAL : classes={total_in_c:.2f} m2 "
                        f"sortie={total_out:.2f} m2 "
                        f"ecart={ecart:+.2f} m2"
                        + (f" (dont {total_autres:.2f} m2 non classes)" if total_autres > 0 else "")),
            })

    return anomalies


def export_tables_to_excel(output_tables, output_path):
    """Écrit le classeur Excel multi-feuilles (une feuille par type de surface).

    Mise en forme appliquée :

    * En-tête projet en italique (9 pt).
    * En-têtes sur-catégories : bleu acier ``#B0C4DE``, fusionnées.
    * En-têtes colonnes : bleu clair ``#DDEEFF``, retour à la ligne automatique.
    * Colonnes sous-total : vert pâle ``#E8F0E8``.
    * Lignes total étage : ``#DCE6F1`` (gras, fusionné Etage + Occupant).
    * Grand total : ``#B8CCE4`` (gras, fusionné).
    * Notes réglementaires en bas, fusionnées sur toute la largeur.

    Args:
        output_tables (dict): Sortie de :func:`Tab_output`.
        output_path (str): Chemin du fichier ``.xlsx`` à créer.
    """

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        for type_su, content in output_tables.items():
            sheet_name = str(type_su)[:31]
            info_lines = content["info"]
            df         = content["data"]
            sc_spans   = content.get("sc_spans", {})

            col_numbers  = content.get("col_numbers", {})
            n_info       = len(info_lines)
            sc_hdr_row   = n_info
            col_hdr_row  = n_info + (1 if sc_spans else 0)
            data_row     = col_hdr_row + 1

            # Créer la feuille vide via un df vide
            pd.DataFrame().to_excel(writer, sheet_name=sheet_name, index=False)
            workbook  = writer.book
            worksheet = writer.sheets[sheet_name]

            cols      = list(df.columns)
            has_occ   = "Occupant" in cols

            # Index cols vs value cols
            index_cols = [c for c in ["Etage", "Occupant"] if c in cols]
            value_cols = [c for c in cols if c not in index_cols]

            # ── Formats ──────────────────────────────────────────────────────
            fmt_info = workbook.add_format({"font_size": 9, "italic": True})

            fmt_col  = workbook.add_format({"bold": True, "bg_color": "#DDEEFF",
                                            "border": 1, "align": "center",
                                            "valign": "vcenter", "text_wrap": True})
            fmt_sc   = workbook.add_format({"bold": True, "bg_color": "#B0C4DE",
                                            "border": 1, "align": "center",
                                            "valign": "vcenter", "font_size": 10})
            fmt_sub  = workbook.add_format({"bold": True, "bg_color": "#E8F0E8",
                                            "border": 1, "align": "center",
                                            "valign": "vcenter", "text_wrap": True})

            # Colonnes calculées SDP (Total TA = orange clair, SDP = bleu acier gras)
            fmt_sdp_ta_hdr = workbook.add_format({"bold": True, "bg_color": "#FCE4D6",
                                                   "border": 1, "align": "center",
                                                   "valign": "vcenter", "text_wrap": True})
            fmt_sdp_hdr    = workbook.add_format({"bold": True, "bg_color": "#B0C4DE",
                                                   "border": 2, "align": "center",
                                                   "valign": "vcenter", "text_wrap": True})
            fmt_sdp_ta_num = workbook.add_format({"bold": True, "bg_color": "#FCE4D6",
                                                   "border": 1, "align": "center",
                                                   "valign": "vcenter",
                                                   "num_format": "#,##0.00"})
            fmt_sdp_num    = workbook.add_format({"bold": True, "bg_color": "#B0C4DE",
                                                   "border": 2, "align": "center",
                                                   "valign": "vcenter",
                                                   "num_format": "#,##0.00"})

            is_sdp = content.get("sdp_fixed", False)

            # Données normales
            fmt_idx  = workbook.add_format({"border": 1})
            fmt_num  = workbook.add_format({"border": 1, "align": "center",
                                            "valign": "vcenter", "text_wrap": True,
                                            "num_format": "#,##0.00"})

            # Total étage (fond rose RGB 255,151,156)
            fmt_etage_merge = workbook.add_format({"bold": True, "bg_color": "#FF979C",
                                                   "border": 1, "align": "left",
                                                   "valign": "vcenter"})
            fmt_etage_num   = workbook.add_format({"bold": True, "bg_color": "#FF979C",
                                                   "border": 1, "align": "center",
                                                   "valign": "vcenter",
                                                   "num_format": "#,##0.00"})

            # Grand total (fond bleu plus foncé)
            fmt_grand_merge = workbook.add_format({"bold": True, "bg_color": "#B8CCE4",
                                                   "border": 1, "align": "left",
                                                   "valign": "vcenter"})
            fmt_grand_num   = workbook.add_format({"bold": True, "bg_color": "#B8CCE4",
                                                   "border": 1, "align": "center",
                                                   "valign": "vcenter",
                                                   "num_format": "#,##0.00"})

            # ── Infos projet ─────────────────────────────────────────────────
            for r, line in enumerate(info_lines):
                worksheet.write(r, 0, line, fmt_info)

            # ── En-têtes sur-catégories ──────────────────────────────────────
            if sc_spans:
                for sc_name, (c_start, c_end) in sc_spans.items():
                    if c_start == c_end:
                        worksheet.write(sc_hdr_row, c_start, sc_name, fmt_sc)
                    else:
                        worksheet.merge_range(sc_hdr_row, c_start,
                                              sc_hdr_row, c_end, sc_name, fmt_sc)

            # ── En-têtes colonnes ────────────────────────────────────────────
            for c_idx, col_name in enumerate(cols):
                num    = col_numbers.get(col_name, "") if col_numbers else ""
                label  = f"{num}\n{col_name}" if num else col_name
                if str(col_name).startswith("[") and str(col_name).endswith("]"):
                    worksheet.write(col_hdr_row, c_idx, "Sous-total", fmt_sub)
                elif is_sdp and col_name == "Total TA":
                    worksheet.write(col_hdr_row, c_idx, label, fmt_sdp_ta_hdr)
                elif is_sdp and col_name == "Total":
                    worksheet.write(col_hdr_row, c_idx, label, fmt_sdp_hdr)
                else:
                    worksheet.write(col_hdr_row, c_idx, label, fmt_col)

            # ── Données ligne par ligne ──────────────────────────────────────
            for row_i, row in df.iterrows():
                xl_row = data_row + row_i
                occupant_val = str(row.get("Occupant", "")) if has_occ else ""
                etage_val    = str(row.get("Etage", ""))

                is_etage_total = occupant_val == "— Total étage"
                is_grand_total = etage_val == "TOTAL"

                if is_grand_total:
                    f_merge, f_num = fmt_grand_merge, fmt_grand_num
                    label = "TOTAL"
                    worksheet.set_row(xl_row, 22)
                    if has_occ:
                        worksheet.merge_range(xl_row, 0, xl_row, 1, label, f_merge)
                    else:
                        worksheet.write(xl_row, 0, label, f_merge)
                    for c_idx, col in enumerate(value_cols):
                        base = len(index_cols)
                        v = row[col]
                        if is_sdp and col == "Total TA":
                            f = fmt_sdp_ta_num
                        elif is_sdp and col == "Total":
                            f = fmt_sdp_num
                        else:
                            f = f_num
                        try:
                            worksheet.write_number(xl_row, base + c_idx, float(v), f)
                        except (TypeError, ValueError):
                            worksheet.write(xl_row, base + c_idx, v if v != "" else 0, f)

                elif is_etage_total:
                    f_merge, f_num = fmt_etage_merge, fmt_etage_num
                    label = f"{etage_val}  —  Total étage"
                    worksheet.set_row(xl_row, 20)
                    if has_occ:
                        worksheet.merge_range(xl_row, 0, xl_row, 1, label, f_merge)
                    else:
                        worksheet.write(xl_row, 0, label, f_merge)
                    for c_idx, col in enumerate(value_cols):
                        base = len(index_cols)
                        v = row[col]
                        if is_sdp and col == "Total TA":
                            f = fmt_sdp_ta_num
                        elif is_sdp and col == "Total":
                            f = fmt_sdp_num
                        else:
                            f = f_num
                        try:
                            worksheet.write_number(xl_row, base + c_idx, float(v), f)
                        except (TypeError, ValueError):
                            worksheet.write(xl_row, base + c_idx, v if v != "" else 0, f)

                else:
                    # Ligne normale — hauteur selon longueur max du texte Occupant
                    occ_len = len(occupant_val)
                    row_h = 30 if occ_len > 30 else 20 if occ_len > 15 else 15
                    worksheet.set_row(xl_row, row_h)
                    for c_idx, col in enumerate(index_cols):
                        worksheet.write(xl_row, c_idx, row[col], fmt_idx)
                    for c_idx, col in enumerate(value_cols):
                        base = len(index_cols)
                        v = row[col]
                        if is_sdp and col == "Total TA":
                            f = fmt_sdp_ta_num
                        elif is_sdp and col == "Total":
                            f = fmt_sdp_num
                        else:
                            f = fmt_num
                        try:
                            worksheet.write_number(xl_row, base + c_idx, float(v), f)
                        except (TypeError, ValueError):
                            worksheet.write(xl_row, base + c_idx, v if v != "" else "", f)

            # ── Largeur des colonnes ─────────────────────────────────────────
            for i, col in enumerate(cols):
                if col == "Etage":
                    width = 8
                elif col == "Occupant":
                    width = 18
                elif col == "Total":
                    width = 12
                else:
                    width = 11
                worksheet.set_column(i, i, width)

            # ── Hauteur en-têtes ─────────────────────────────────────────────
            max_cat_len = max((len(str(c)) for c in cols if c not in ("Etage", "Occupant", "Total")), default=10)
            hdr_h = max(30, min(60, (max_cat_len // 11 + 1) * 15))
            worksheet.set_row(col_hdr_row, hdr_h)

            # ── Nota en bas de tableau ───────────────────────────────────────
            nota     = content.get("nota", [])
            mesurage = content.get("mesurage", "")
            n_cols   = len(cols)

            nota_intro_line = (
                f"Nota : les superficies ont été calculées après mesurage des locaux en {mesurage}. "
                "\n Les désignations ont été déterminées en fonction des signes apparents constatés "
                "le jour du mesurage."
            )

            nota_start = data_row + len(df) + 1
            fmt_sep = workbook.add_format({"bottom": 1, "bottom_color": "#AAAAAA"})
            worksheet.write(nota_start - 1, 0, "", fmt_sep)
            fmt_nota_intro = workbook.add_format({
                "font_size": 9, "italic": True, "text_wrap": True, "valign": "top",
            })
            fmt_nota = workbook.add_format({
                "font_size": 8, "italic": True, "fg_color": "#555555",
                "text_wrap": True, "valign": "top",
            })

            # Ligne intro mesurage (toujours présente)
            if n_cols > 1:
                worksheet.merge_range(nota_start, 0, nota_start, n_cols - 1,
                                      nota_intro_line, fmt_nota_intro)
            else:
                worksheet.write(nota_start, 0, nota_intro_line, fmt_nota_intro)
            approx = max(1, 11 * n_cols)
            worksheet.set_row(nota_start, max(1, len(nota_intro_line) // approx + 1) * 12 + 4)

            # Nota réglementaires
            for i, line in enumerate(nota):
                r = nota_start + 1 + i
                if n_cols > 1:
                    worksheet.merge_range(r, 0, r, n_cols - 1, line, fmt_nota)
                else:
                    worksheet.write(r, 0, line, fmt_nota)
                approx_chars_per_line = max(1, 11 * n_cols)
                nb_lines = max(1, len(line) // approx_chars_per_line + 1)
                worksheet.set_row(r, nb_lines * 12 + 4)


def html_from_excel(xlsx_path, infos):
    """Relit un fichier ``.xlsx`` généré par Rfill et reconstruit ``output_tables`` pour re-générer le HTML.

    Chaque feuille de l'Excel correspond à un type de surface (SUB, SU, SHO…).
    Les lignes d'en-tête projet (italique, avant le tableau) sont ignorées : seul le
    tableau de données est relu, les métadonnées proviennent de ``infos``.

    Args:
        xlsx_path (str): Chemin du ``.xlsx`` produit par :func:`export_tables_to_excel`.
        infos (dict): Métadonnées projet (batiment, adresse, proprio, cadastre,
            date, dossier, mesurage).

    Returns:
        dict: ``output_tables`` au même format que :func:`Tab_output`,
        prêt pour :func:`export_tables_to_html`.

    Raises:
        ValueError: Si le fichier ne contient aucune feuille reconnue.
    """
    import openpyxl

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    output_tables = {}

    for sheet_name in wb.sheetnames:
        type_su = sheet_name.strip().upper()

        # Lire toutes les lignes de la feuille
        ws = wb[sheet_name]
        rows = list(ws.values)
        if not rows:
            continue

        # Trouver la ligne d'en-tête du tableau : première ligne dont la première
        # cellule vaut "Etage" (les lignes d'infos projet sont avant)
        header_idx = None
        for i, row in enumerate(rows):
            if row and str(row[0]).strip() == "Etage":
                header_idx = i
                break

        if header_idx is None:
            continue

        headers = [str(c).strip() if c is not None else "" for c in rows[header_idx]]

        # Lignes de données : tout ce qui suit l'en-tête jusqu'aux lignes nota
        # (nota : ligne dont toutes les colonnes sauf la première sont vides ou None)
        data_rows = []
        for row in rows[header_idx + 1:]:
            if row[0] is None:
                break
            vals = [c for c in row[1:] if c is not None and str(c).strip() not in ("", "None")]
            if not vals and str(row[0]).strip().startswith("Nota"):
                break
            data_rows.append(row)

        if not data_rows:
            continue

        df = pd.DataFrame(data_rows, columns=headers)

        # Convertir colonnes numériques
        index_cols = [c for c in ("Etage", "Occupant") if c in df.columns]
        for col in df.columns:
            if col not in index_cols:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna("")

        nota = glossary_surf.nota_surf.get(type_su, [])
        is_sdp = type_su == "SDP"

        output_tables[type_su] = {
            "info": _build_info_lines(type_su, infos),
            "data": df,
            "sc_spans": {},
            "nota": nota,
            "mesurage": infos.get("mesurage", "").strip(),
            "sdp_fixed": is_sdp,
            "col_numbers": _SDP_COL_NUMBERS if is_sdp else {},
        }

    if not output_tables:
        raise ValueError("Aucune feuille reconnue dans le fichier Excel.")

    return output_tables


def update_glossary(mapping_df, type_su, super_cats_for_type=None):
    """Persiste les affectations classifiées et la structure des catégories dans ``glossary_surf.py``.

    Ajoute les affectations dont ``cat != "autres"`` si absentes du glossaire existant.
    Si ``super_cats_for_type`` est fourni, met à jour également ``superficie_names`` et
    ``predefined_cats`` pour persister la structure des super-catégories et catégories.
    L'écriture est **atomique** : fichier temporaire → :func:`shutil.move` pour éviter
    toute corruption en cas d'interruption. La syntaxe est validée par :func:`ast.parse`
    avant l'écriture. Le module est rechargé via :func:`_reload_glossary` après modification.

    Args:
        mapping_df (pandas.DataFrame): Table ``['Affectation', 'cat']``
            issue de l'interface utilisateur.
        type_su (str): Code de surface (ex. ``"SU"``).
        super_cats_for_type (dict, optional): Structure ``{super_cat_name: [cat_name, ...]}``
            contenant les super-catégories et leurs catégories associées. Si fourni,
            met à jour ``superficie_names[type_su]`` et ``predefined_cats[type_su]``.

    Returns:
        bool: ``True`` si le fichier a été modifié, ``False`` sinon.
    """
    type_su_glo = glossary_surf.denom_surf.get(type_su)
    if not type_su_glo:
        return False

    type_glossary = glossary_surf.glossary_surf[type_su_glo]
    modified = False

    for _, row in mapping_df.iterrows():
        aff = str(row["Affectation"]).strip()
        cat = row["cat"]
        if cat == "autres" or not aff:
            continue
        if cat not in type_glossary:
            type_glossary[cat] = []
            modified = True
        existing_lower = [k.lower() for k in type_glossary[cat]]
        if aff.lower() not in existing_lower:
            type_glossary[cat].append(aff.lower())
            modified = True

    if super_cats_for_type:
        new_sc_names = list(super_cats_for_type.keys())
        new_cats = [c for cats in super_cats_for_type.values() for c in cats]

        # S'assurer que toutes les catégories ont une entrée dans le glossaire (même vide)
        for cat in new_cats:
            if cat and cat not in type_glossary:
                type_glossary[cat] = []
                modified = True

        if glossary_surf.superficie_names.get(type_su) != new_sc_names:
            glossary_surf.superficie_names[type_su] = new_sc_names
            modified = True

        if glossary_surf.predefined_cats.get(type_su) != new_cats:
            glossary_surf.predefined_cats[type_su] = new_cats
            modified = True

    if not modified:
        return False

    content = (
        "glossary_surf = "
        + pprint.pformat(glossary_surf.glossary_surf, width=120, sort_dicts=False)
        + "\n\ndenom_surf = "
        + pprint.pformat(glossary_surf.denom_surf, width=120, sort_dicts=False)
        + "\n\nreal_su_name = "
        + pprint.pformat(glossary_surf.real_su_name, width=120, sort_dicts=False)
        + "\n\nsuperficie_names = "
        + pprint.pformat(glossary_surf.superficie_names, width=120, sort_dicts=False)
        + "\n\npredefined_cats = "
        + pprint.pformat(glossary_surf.predefined_cats, width=120, sort_dicts=False)
        + "\n\ncat_colors = "
        + pprint.pformat(glossary_surf.cat_colors, width=120, sort_dicts=False)
        + "\n\nnota_surf = "
        + pprint.pformat(glossary_surf.nota_surf, width=120, sort_dicts=False)
        + "\n"
    )

    # Valider la syntaxe avant d'écrire
    ast.parse(content)

    # Écriture atomique : temp → rename (évite la corruption si interruption)
    file_path = _glossary_path()
    fd, tmp_path = tempfile.mkstemp(suffix=".py", dir=os.path.dirname(file_path))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(content)
        shutil.move(tmp_path, file_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise

    _reload_glossary()
    return True


# =========================================================
#  EXPORT HTML (A4 paysage imprimable)
# =========================================================

def _img_dir():
    """Retourne le chemin du dossier ``img/`` adapté au mode d'exécution.

    En mode *frozen* PyInstaller (onefile), cherche d'abord dans ``sys._MEIPASS``
    (dossier d'extraction temporaire), puis à côté de l'exécutable.

    Returns:
        str: Chemin absolu vers le dossier ``img/``.
    """
    if getattr(sys, "frozen", False):
        # 1) dossier extraction PyInstaller (onefile), 2) à côté de l'exe
        base = getattr(sys, "_MEIPASS", None) or os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "img")


def _img_to_base64(filename):
    """Encode une image du dossier ``img/`` en data URI base64.

    Args:
        filename (str): Nom du fichier image (ex. ``"logo_ge.jpg"``).

    Returns:
        str: Data URI ``data:image/<mime>;base64,...``, ou ``""`` si le fichier
        est absent.
    """
    path = os.path.join(_img_dir(), filename)
    if not os.path.isfile(path):
        return ""
    ext = os.path.splitext(filename)[1].lower().lstrip(".")
    mime = "jpeg" if ext in ("jpg", "jpeg") else ext
    with open(path, "rb") as f:
        data = base64.b64encode(f.read()).decode("ascii")
    return f"data:image/{mime};base64,{data}"


def _fmt_num(v):
    """Formate une valeur numérique pour l'affichage HTML.

    Séparateur des milliers : espace insécable ; séparateur décimal : virgule.

    Args:
        v: Valeur à formater (numérique, chaîne ou ``None``).

    Returns:
        str: Chaîne formatée (ex. ``"1 234,56"``), chaîne échappée si non
        numérique, ou ``""`` si *v* est ``None`` ou vide.
    """
    try:
        f = float(v)
    except (TypeError, ValueError):
        return html.escape(str(v)) if v not in (None, "") else ""
    return f"{f:,.2f}".replace(",", " ").replace(".", ",")


def export_tables_to_html(output_tables, infos, html_path):
    """Génère un rapport HTML auto-contenu, imprimable A4 paysage.

    Produit un fichier ``.html`` unique (logos et tampon encodés en base64,
    aucune dépendance externe) avec :

    * Une ``<section class="page">`` par type de surface.
    * En-tête : logos gauche · infos projet centrées · date/dossier droite.
    * Navigation par onglets JavaScript (masquée à l'impression via
      ``@media print``).
    * Pied de page : nota intro + notes réglementaires + tampon.
    * CSS ``@page { size: A4 landscape; }`` pour l'impression directe.

    Args:
        output_tables (dict): Sortie de :func:`Tab_output`.
        infos (dict): Métadonnées projet — clés : ``batiment``, ``adresse``,
            ``proprio``, ``cadastre``, ``date``, ``dossier``, ``mesurage``.
            Le champ ``mesurage`` remplit la phrase
            *"superficies calculées après mesurage en …"*.
        html_path (str): Chemin du fichier ``.html`` à créer.
    """
    logo_ge     = _img_to_base64("logo_ge.jpg")
    logo_rtaxes = _img_to_base64("logo_rtaxes.jpg")
    tampon      = _img_to_base64("Tampon GE pr tblx.jpg")

    esc = html.escape
    mesurage = infos.get("mesurage", "").strip() or "xx 2026"
    date_v   = esc(infos.get("date", "").strip())
    dossier  = esc(infos.get("dossier", "").strip())
    batiment = esc(infos.get("batiment", "").strip())
    adresse  = esc(infos.get("adresse", "").strip())
    proprio  = esc(infos.get("proprio", "").strip())
    cadastre = esc(infos.get("cadastre", "").strip())

    css = """
@page { size: A4 landscape; margin: 10mm 18mm; }
* { box-sizing: border-box; }
body { font-family: Arial, sans-serif; font-size: 10px; color: #222; margin: 0; }
.page { page-break-after: always; display: flex; flex-direction: column;
        min-height: 190mm; padding: 0 4mm; }
.page:last-child { page-break-after: auto; }
header { display: grid; grid-template-columns: 1fr 2fr 1fr; align-items: center;
         border-bottom: 2px solid #1a2a4a; padding-bottom: 6px; margin-bottom: 10px; gap: 8px; }
header .logos { display: flex; align-items: center; gap: 8px; }
header .logos img { max-height: 70px; max-width: 120px; object-fit: contain; }
header .logos .cabinet-info { font-size: 8px; line-height: 1.6; color: #333; }
header .infos { text-align: center; font-size: 9px; line-height: 1.35; }
header .infos .title { font-size: 12px; font-weight: bold; color: #1a2a4a;
                       text-transform: uppercase; margin-bottom: 3px; }
header .meta { text-align: right; font-size: 9px; line-height: 1.4; }
header .meta .row { margin-bottom: 2px; }
header .meta .k { color: #555; font-weight: bold; margin-right: 4px; }
header .stamp-block { display: flex; flex-direction: column; align-items: center; gap: 4px; }
header .stamp-block img.tampon { max-height: 70px; max-width: 160px; object-fit: contain; }
header .stamp-block img.oge { max-height: 28px; max-width: 120px; object-fit: contain; }
table.surf { width: 100%; border-collapse: collapse; font-size: 9px;
             table-layout: fixed; }
table.surf th, table.surf td { border: 1px solid #888; padding: 3px 4px;
                                text-align: center; vertical-align: middle;
                                overflow-wrap: break-word; word-break: normal;
                                print-color-adjust: exact; -webkit-print-color-adjust: exact; }
table.surf th { background: #B0ACC4; color: #111; font-weight: bold; }
table.surf th.sc { background: #B0ACC4; }
table.surf th.sub, table.surf th.sdp-total { background: #FF979C; }
table.surf th span.colnum { display: block; font-size: 8px; color: #333; font-weight: normal; margin-bottom: 2px; }
table.surf td.idx { text-align: left; }
table.surf td.etage-cell { background: #B0ACC4; font-weight: bold; }
table.surf td.num { text-align: center; }
table.surf td.subtotal, table.surf td.sdp-total { background: #FF979C !important; font-weight: bold; text-align: center; }
table.surf tr.total-etage td { background: #FF979C; font-weight: bold; text-align: left; }
table.surf tr.total-etage td.num { text-align: center; }
table.surf tr.total td { background: #FF979C; font-weight: bold; text-align: left; }
table.surf tr.total td.num { text-align: center; }
.footer { display: flex; align-items: flex-start; gap: 12px; margin-top: 6px; }
.footer .notas { flex: 1; min-width: 0; }
.nota-intro { font-size: 9px; font-style: italic; color: #333;
              line-height: 1.5; border-left: 2px solid #b0c4de; padding-left: 6px; }
.nota { margin-top: 4px; font-size: 8px; color: #555; line-height: 1.4; }
.nota p { margin: 2px 0; }
.stamp { flex: 0 0 auto; display: flex; flex-direction: column; align-items: center;
         justify-content: flex-start; gap: 4px; padding-left: 8px; }
.stamp img.tampon { max-height: 70px; max-width: 160px; object-fit: contain; }
.stamp img.oge { max-height: 70px; max-width: 160px; object-fit: contain; }
/* Onglets (écran uniquement, masqués à l'impression) */
@media screen {
  body { background: #eef2f7; padding: 10px 0; }
  .tabs { position: sticky; top: 0; z-index: 10; background: #1a2a4a;
          padding: 8px 16px; display: flex; gap: 4px; flex-wrap: wrap;
          box-shadow: 0 2px 6px rgba(0,0,0,.15); margin-bottom: 12px; }
  .tabs a { color: #cfd8e8; text-decoration: none; padding: 6px 14px;
            border-radius: 4px 4px 0 0; font-size: 11px; font-weight: bold;
            letter-spacing: 0.3px; }
  .tabs a:hover { background: #2a3a5a; color: #fff; }
  .tabs a.active { background: #fff; color: #1a2a4a; }
  .tabs .btn-print { margin-left: auto; background: #2e7d32; color: #fff;
                     border: none; padding: 6px 16px; border-radius: 4px;
                     font-size: 11px; font-weight: bold; cursor: pointer;
                     letter-spacing: 0.3px; }
  .tabs .btn-print:hover { background: #1b5e20; }
  .page { background: #fff; width: 297mm; margin: 0 auto 14px auto;
          padding: 10mm 18mm; box-shadow: 0 1px 4px rgba(0,0,0,.12); }
  body:not([data-all]) .page { display: none; }
  body:not([data-all]) .page.active { display: flex; }
}
@media print { .tabs { display: none !important; } }
"""

    _CABINET_INFO = (
        "66-68 rue du Bocage<br>"
        "37540 SAINT-CYR-SUR-LOIRE<br>"
        "&#9990;&nbsp;02 47 42 14 98<br>"
        "&#9993;&nbsp;benoit.decorbier@geometre-expert.fr"
    )

    def render_header(sheet_title):
        # Gauche : logo cabinet + infos cabinet
        cabinet_html = ""
        if logo_rtaxes:
            cabinet_html += f'<img src="{logo_rtaxes}" alt="">'
        cabinet_html += f'<div class="cabinet-info">{_CABINET_INFO}</div>'

        # Ligne 1 : Bâtiment / Adresse
        bat_adr_parts = []
        if batiment:
            bat_adr_parts.append(f"<b>Bâtiment :</b> {batiment}")
        if adresse:
            bat_adr_parts.append(f"<b>Adresse :</b> {adresse}")
        bat_adr_html = " &nbsp;·&nbsp; ".join(bat_adr_parts)

        # Ligne 2 : Propriétaire / Cadastre
        prop_cad_parts = []
        if proprio:
            prop_cad_parts.append(f"<b>Propriétaire :</b> {proprio}")
        if cadastre:
            prop_cad_parts.append(f"<b>Cadastre :</b> {cadastre}")
        prop_cad_html = " &nbsp;·&nbsp; ".join(prop_cad_parts)

        infos_lines = "".join(
            f"<div>{line}</div>" for line in [bat_adr_html, prop_cad_html] if line
        )

        meta_date_dossier = "".join([
            f'<div class="row"><span class="k">Date :</span>{date_v}</div>' if date_v else "",
            f'<div class="row"><span class="k">Dossier :</span>{dossier}</div>' if dossier else "",
        ])

        return f"""
<header>
  <div class="logos">{cabinet_html}</div>
  <div class="infos">
    <div class="title">{esc(sheet_title)}</div>
    {infos_lines}
    <div style="font-size:9px;margin-top:3px">{meta_date_dossier}</div>
  </div>
  <div class="stamp-block"></div>
</header>
"""

    def render_table(df, sc_spans, col_numbers=None, sdp_fixed=False):
        cols = list(df.columns)
        index_cols = [c for c in ("Etage", "Occupant") if c in cols]
        value_cols = [c for c in cols if c not in index_cols]
        n_cols = len(cols)
        has_occupant = "Occupant" in index_cols

        # Colonne "Total" SDP = dernière colonne → rose
        sdp_total_col = "Total" if sdp_fixed else None

        # Largeur colonne Etage : basée sur la valeur la plus longue du tableau
        etage_vals = [str(r) for r in df.get("Etage", pd.Series([])) if str(r) not in ("TOTAL",)]
        max_etage_len = max((len(v) for v in etage_vals), default=10)
        etage_width = f"{max(55, min(max_etage_len * 6, 110))}px"

        col_widths = []
        for c in cols:
            if c == "Etage":
                col_widths.append(etage_width)
            elif c == "Occupant":
                col_widths.append("90px")
            elif c in ("Total", "Total TA"):
                col_widths.append("52px")
            elif str(c).startswith("[") and str(c).endswith("]"):
                col_widths.append("52px")
            else:
                col_widths.append("auto")
        colgroup = "<colgroup>" + "".join(f'<col style="width:{w}">' for w in col_widths) + "</colgroup>"

        # Ligne sur-catégories si présentes
        sc_row = ""
        if sc_spans:
            covered = {}
            occupied = set()
            for sc_name, (c_start, c_end) in sc_spans.items():
                covered[c_start] = (sc_name, c_end - c_start + 1)
                for i in range(c_start, c_end + 1):
                    occupied.add(i)
            cells = []
            i = 0
            while i < n_cols:
                if i in covered:
                    name, span = covered[i]
                    cells.append(f'<th class="sc" colspan="{span}">{esc(name)}</th>')
                    i += span
                elif i in occupied:
                    i += 1
                elif cols[i] in index_cols:
                    # Etage / Occupant : rowspan=2 pour fusionner avec la ligne sc
                    cells.append(f'<th class="sc" rowspan="2">{esc(str(cols[i]))}</th>')
                    i += 1
                else:
                    cells.append('<th class="sc"></th>')
                    i += 1
            sc_row = "<tr>" + "".join(cells) + "</tr>"

        # Ligne en-tête colonnes (on saute les index_cols déjà fusionnées dans sc_row)
        hdr_cells = []
        for c in cols:
            if sc_row and c in index_cols:
                continue  # déjà émis avec rowspan=2 dans sc_row
            cs  = str(c)
            num = col_numbers.get(c, "") if col_numbers else ""
            is_sdp_tot = (c == sdp_total_col)
            extra_cls = " sdp-total" if is_sdp_tot else ""
            if cs.startswith("[") and cs.endswith("]"):
                hdr_cells.append(f'<th class="sub{extra_cls}">Sous-total</th>')
            elif num:
                hdr_cells.append(f'<th class="{extra_cls.strip()}"><span class="colnum">{esc(num)}</span><br>{esc(cs)}</th>')
            else:
                hdr_cells.append(f'<th class="{extra_cls.strip()}">{esc(cs)}</th>')
        hdr_row = "<tr>" + "".join(hdr_cells) + "</tr>"

        # Lignes de données
        body_rows = []
        for _, row in df.iterrows():
            occupant_val   = str(row.get("Occupant", "")) if has_occupant else ""
            etage_val      = str(row.get("Etage", ""))
            is_total_etage = (occupant_val == "— Total étage") or etage_val.endswith("— Total étage")
            is_grand       = etage_val == "TOTAL"

            if is_grand:
                tds = []
                if has_occupant:
                    tds.append(f'<td colspan="{len(index_cols)}" class="idx">TOTAL</td>')
                else:
                    tds.append('<td class="idx">TOTAL</td>')
                for c in value_cols:
                    extra = " sdp-total" if c == sdp_total_col else ""
                    tds.append(f'<td class="num{extra}">{_fmt_num(row[c])}</td>')
                body_rows.append(f'<tr class="total">{"".join(tds)}</tr>')
                continue

            if is_total_etage:
                label_te = occupant_val if has_occupant else etage_val
                tds = [f'<td colspan="{len(index_cols)}" class="idx">{esc(label_te)}</td>']
                for c in value_cols:
                    extra = " sdp-total" if c == sdp_total_col else ""
                    tds.append(f'<td class="num{extra}">{_fmt_num(row[c])}</td>')
                body_rows.append(f'<tr class="total-etage">{"".join(tds)}</tr>')
                continue

            tds = []
            if has_occupant:
                tds.append(f'<td class="idx etage-cell">{esc(etage_val)}</td>')
                tds.append(f'<td class="idx">{esc(occupant_val)}</td>')
            else:
                tds.append(f'<td class="idx etage-cell">{esc(etage_val)}</td>')

            for c in value_cols:
                is_sub = str(c).startswith("[") and str(c).endswith("]")
                extra = " sdp-total" if c == sdp_total_col else (" subtotal" if is_sub else "")
                tds.append(f'<td class="num{extra}">{_fmt_num(row[c])}</td>')
            body_rows.append(f"<tr>{''.join(tds)}</tr>")

        return (
            '<table class="surf">'
            + colgroup
            + "<thead>" + sc_row + hdr_row + "</thead>"
            + "<tbody>" + "".join(body_rows) + "</tbody>"
            + "</table>"
        )

    nota_intro = (
        f"Nota : les superficies ont été calculées après mesurage des locaux en {esc(mesurage)}.<br>"
        "Les désignations ont été déterminées en fonction des signes apparents constatés "
        "le jour du mesurage."
    )

    stamp_parts = ""
    if tampon:
        stamp_parts += f'<img class="tampon" src="{tampon}" alt="">'
    if logo_ge:
        stamp_parts += f'<img class="oge" src="{logo_ge}" alt="">'
    stamp_html = f'<div class="stamp">{stamp_parts}</div>' if stamp_parts else ""

    pages = []
    tabs = []
    for idx, (type_su, content) in enumerate(output_tables.items()):
        sheet_title = glossary_surf.real_su_name.get(type_su, type_su)
        df          = content["data"]
        sc_spans    = content.get("sc_spans", {})
        nota        = content.get("nota", [])
        col_numbers = content.get("col_numbers", {})

        nota_list_html = ""
        if nota:
            lines = "".join(f"<p>{esc(line)}</p>" for line in nota)
            nota_list_html = f'<div class="nota">{lines}</div>'

        footer_html = (
            '<div class="footer">'
            '<div class="notas">'
            + f'<div class="nota-intro">{nota_intro}</div>'
            + nota_list_html
            + "</div>"
            + stamp_html
            + "</div>"
        )

        page_id   = f"page-{esc(str(type_su))}"
        active    = " active" if idx == 0 else ""
        tabs.append(
            f'<a href="#{page_id}" class="tab-link{active}" '
            f'data-target="{page_id}">{esc(str(type_su))}</a>'
        )
        page_html = (
            f'<section class="page{active}" id="{page_id}">'
            + render_header(sheet_title)
            + render_table(df, sc_spans, col_numbers, sdp_fixed=content.get("sdp_fixed", False))
            + footer_html
            + "</section>"
        )
        pages.append(page_html)

    tabs_html = (
        '<nav class="tabs">'
        + "".join(tabs)
        + '<a href="#" class="tab-link" data-target="__all__">Tout afficher</a>'
        + '<button class="btn-print" onclick="window.print()">Imprimer / PDF</button>'
        + "</nav>"
    )

    js = """
<script>
document.querySelectorAll('.tab-link').forEach(function(a){
  a.addEventListener('click', function(e){
    e.preventDefault();
    var t = this.dataset.target;
    document.querySelectorAll('.tab-link').forEach(function(x){x.classList.remove('active');});
    this.classList.add('active');
    if (t === '__all__') {
      document.body.setAttribute('data-all','1');
    } else {
      document.body.removeAttribute('data-all');
      document.querySelectorAll('.page').forEach(function(p){
        p.classList.toggle('active', p.id === t);
      });
      document.getElementById(t).scrollIntoView({behavior:'smooth', block:'start'});
    }
  });
});
</script>
"""

    doc = (
        "<!DOCTYPE html>\n"
        '<html lang="fr"><head><meta charset="utf-8">'
        f"<title>Rapport de surfaces — {dossier or batiment}</title>"
        f"<style>{css}</style></head><body>"
        + tabs_html
        + "".join(pages)
        + js
        + "</body></html>"
    )

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(doc)