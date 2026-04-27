import ast
import os
import pprint
import shutil
import tempfile
import importlib

import pandas as pd
import glossary_surf

# =========================================================
#  LOGIQUE METIER
# =========================================================

def extract_info(file_path):
    df = pd.read_excel(file_path, header=None)

    # Les lignes de données ont un entier en colonne 0 (numéro de séquence AutoCAD)
    df[0] = pd.to_numeric(df[0], errors="coerce")
    df = df[df[0].notna()].copy().reset_index(drop=True)

    df.columns = ["num_piece", "Descriptif", "Calque", "N", "Affectation",
                  "Occupant", "Aire", "Etage", "Chambre", "Lot", "Accessibilité aux publics"]

    # Type de surface directement depuis le calque : "SUB Contours" → "SUB"
    df["type_su"] = df["Calque"].astype(str).str.strip().str.split().str[0]

    return df


def tab_cd_type(df):
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
    """
    Construit un mapping Affectation → Catégorie basé sur le glossaire,
    mais uniquement pour les affectations présentes dans df_t.
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
    """
    df_t : DataFrame avec colonnes ['Etage', 'Affectation', 'Aire', 'type_su']
    mapping_df : DataFrame ['Affectation', 'cat'] modifié via l'interface.
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



def Tab_output(df_tcd, infos, super_cat_map=None):
    """
    Transforme df_tcd (Etage, Occupant, cat, Aire, type_su)
    en tableau Excel final avec en-tête projet + pivot par catégorie.
    """
    output_tables = {}
    type_su_list = pd.unique(df_tcd["type_su"].values)

    for t in type_su_list:
        df_t = df_tcd[df_tcd["type_su"] == t].copy()
        df_t = df_t[df_t["cat"] != "autres"]

        has_occupant = (
            t != "SU"
            and "Occupant" in df_t.columns
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

        cat_cols = [c for c in table.columns if c not in index_cols]
        table["Total"] = table[cat_cols].sum(axis=1)

        # Sous-totaux par sur-catégorie
        sc_map = (super_cat_map or {}).get(t, {})
        sc_spans = {}  # {sc_name: (col_start_idx, col_end_idx)} dans le df final
        if sc_map:
            ordered = list(index_cols)
            for sc_name, sc_cats in sc_map.items():
                sc_actual = [c for c in sc_cats if c in cat_cols]
                if sc_actual:
                    ordered.extend(sc_actual)
                    sub_col = f"[{sc_name}]"
                    table[sub_col] = table[sc_actual].sum(axis=1)
                    ordered.append(sub_col)
            remaining = [c for c in cat_cols if c not in ordered]
            ordered.extend(remaining)
            ordered.append("Total")
            table = table[[c for c in ordered if c in table.columns]]
            # Calculer les spans de colonnes pour l'export Excel
            final_cols = list(table.columns)
            for sc_name, sc_cats in sc_map.items():
                sc_actual = [c for c in sc_cats if c in cat_cols]
                if sc_actual:
                    sub_col = f"[{sc_name}]"
                    sc_spans[sc_name] = (final_cols.index(sc_actual[0]),
                                         final_cols.index(sub_col))

        all_value_cols = [c for c in table.columns if c not in index_cols]

        # Sous-totaux par étage (seulement quand plusieurs occupants)
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
        for c in all_value_cols:
            total_row[c] = pd.to_numeric(
                table.loc[table.get("Occupant", pd.Series([""] * len(table))) != "— Total étage", c],
                errors="coerce"
            ).sum()
        table = pd.concat([table, pd.DataFrame([total_row])], ignore_index=True)

        info_lines = [
            glossary_surf.real_su_name.get(t, t),
            f"Bâtiment : {infos['batiment']}",
            f"Adresse : {infos['adresse']}",
            f"Propriétaire : {infos['proprio']}",
            f"Cadastre : {infos['cadastre']}",
            f"Date : {infos['date']}",
            f"Dossier : {infos['dossier']}",
            f"Mesurage : {infos['mesurage']}",
            "",
        ]

        output_tables[t] = {"info": info_lines, "data": table, "sc_spans": sc_spans}

    return output_tables


def export_tables_to_excel(output_tables, output_path):
    """
    output_tables : dict {type_su: {"info": [...], "data": DataFrame}}
    output_path : chemin du fichier .xlsx à créer
    """

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        for type_su, content in output_tables.items():
            sheet_name = str(type_su)[:31]
            info_lines = content["info"]
            df         = content["data"]
            sc_spans   = content.get("sc_spans", {})

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

            # Données normales
            fmt_idx  = workbook.add_format({"border": 1})
            fmt_num  = workbook.add_format({"border": 1, "align": "center",
                                            "valign": "vcenter", "text_wrap": True,
                                            "num_format": "#,##0.00"})

            # Total étage (fond bleu clair, merge Etage+Occupant)
            fmt_etage_merge = workbook.add_format({"bold": True, "bg_color": "#DCE6F1",
                                                   "border": 1, "align": "left",
                                                   "valign": "vcenter"})
            fmt_etage_num   = workbook.add_format({"bold": True, "bg_color": "#DCE6F1",
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
                if str(col_name).startswith("[") and str(col_name).endswith("]"):
                    worksheet.write(col_hdr_row, c_idx, "Sous-total", fmt_sub)
                else:
                    worksheet.write(col_hdr_row, c_idx, col_name, fmt_col)

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
                        try:
                            worksheet.write_number(xl_row, base + c_idx, float(v), f_num)
                        except (TypeError, ValueError):
                            worksheet.write(xl_row, base + c_idx, v if v != "" else 0, f_num)

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
                        try:
                            worksheet.write_number(xl_row, base + c_idx, float(v), f_num)
                        except (TypeError, ValueError):
                            worksheet.write(xl_row, base + c_idx, v if v != "" else 0, f_num)

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
                        try:
                            worksheet.write_number(xl_row, base + c_idx, float(v), fmt_num)
                        except (TypeError, ValueError):
                            worksheet.write(xl_row, base + c_idx, v if v != "" else "", fmt_num)

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
            # Hauteur calculée selon la longueur max des noms de catégorie (width=11)
            max_cat_len = max((len(str(c)) for c in cols if c not in ("Etage", "Occupant", "Total")), default=10)
            hdr_h = max(30, min(60, (max_cat_len // 11 + 1) * 15))
            worksheet.set_row(col_hdr_row, hdr_h)


def update_glossary(mapping_df, type_su):
    """
    Ajoute dans glossary_surf.py les affectations classifiées (hors 'autres')
    qui n'y figurent pas encore, puis recharge le module.
    Retourne True si le fichier a été modifié.
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
        existing_lower = [k.lower() for k in type_glossary[cat]]
        if aff.lower() not in existing_lower:
            type_glossary[cat].append(aff.lower())
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
        + "\n"
    )

    # Valider la syntaxe avant d'écrire
    ast.parse(content)

    # Écriture atomique : temp → rename (évite la corruption si interruption)
    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "glossary_surf.py")
    fd, tmp_path = tempfile.mkstemp(suffix=".py", dir=os.path.dirname(file_path))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(content)
        shutil.move(tmp_path, file_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise

    importlib.reload(glossary_surf)
    return True