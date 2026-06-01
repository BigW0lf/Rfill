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


def _glossary_path():
    """Chemin du fichier glossary_surf.py — adapté au mode frozen (exe PyInstaller)."""
    if getattr(sys, 'frozen', False):
        # En mode exe : dossier à côté de Rfill.exe
        return os.path.join(os.path.dirname(sys.executable), "glossary_surf.py")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "glossary_surf.py")


def _reload_glossary():
    """Recharge le module glossary_surf depuis le fichier sur disque."""
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

        nota = glossary_surf.nota_surf.get(t, [])
        output_tables[t] = {"info": info_lines, "data": table, "sc_spans": sc_spans, "nota": nota}

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
            max_cat_len = max((len(str(c)) for c in cols if c not in ("Etage", "Occupant", "Total")), default=10)
            hdr_h = max(30, min(60, (max_cat_len // 11 + 1) * 15))
            worksheet.set_row(col_hdr_row, hdr_h)

            # ── Nota en bas de tableau ───────────────────────────────────────
            nota = content.get("nota", [])
            if nota:
                nota_start = data_row + len(df) + 1
                fmt_sep = workbook.add_format({"bottom": 1, "bottom_color": "#AAAAAA"})
                worksheet.write(nota_start - 1, 0, "", fmt_sep)
                fmt_nota = workbook.add_format({
                    "font_size": 8, "italic": True, "fg_color": "#555555",
                    "text_wrap": True, "valign": "top",
                })
                n_cols = len(cols)
                for i, line in enumerate(nota):
                    r = nota_start + i
                    if n_cols > 1:
                        worksheet.merge_range(r, 0, r, n_cols - 1, line, fmt_nota)
                    else:
                        worksheet.write(r, 0, line, fmt_nota)
                    # Hauteur proportionnelle à la longueur du texte (col width ≈ 11*n_cols chars)
                    approx_chars_per_line = max(1, 11 * n_cols)
                    nb_lines = max(1, len(line) // approx_chars_per_line + 1)
                    worksheet.set_row(r, nb_lines * 12 + 4)


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
        + "\n\nsuperficie_names = "
        + pprint.pformat(glossary_surf.superficie_names, width=120, sort_dicts=False)
        + "\n\npredefined_cats = "
        + pprint.pformat(glossary_surf.predefined_cats, width=120, sort_dicts=False)
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
    """Dossier des images — adapté au mode frozen (PyInstaller onefile)."""
    if getattr(sys, "frozen", False):
        # 1) dossier extraction PyInstaller (onefile), 2) à côté de l'exe
        base = getattr(sys, "_MEIPASS", None) or os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "img")


def _img_to_base64(filename):
    """Encode une image du dossier img/ en data URI base64. Retourne '' si absente."""
    path = os.path.join(_img_dir(), filename)
    if not os.path.isfile(path):
        return ""
    ext = os.path.splitext(filename)[1].lower().lstrip(".")
    mime = "jpeg" if ext in ("jpg", "jpeg") else ext
    with open(path, "rb") as f:
        data = base64.b64encode(f.read()).decode("ascii")
    return f"data:image/{mime};base64,{data}"


def _fmt_num(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return html.escape(str(v)) if v not in (None, "") else ""
    return f"{f:,.2f}".replace(",", " ").replace(".", ",")


def export_tables_to_html(output_tables, infos, html_path):
    """
    Génère un rapport HTML imprimable A4 paysage, une page par type de surface.
    output_tables : sortie de Tab_output()
    infos : dict (batiment, adresse, proprio, cadastre, date, dossier, mesurage)
    """
    logo_ge     = _img_to_base64("logo_ge.jpg")
    logo_rtaxes = _img_to_base64("logo_rtaxes.jpg")
    tampon      = _img_to_base64("Tampom_rtaxes.jpg")

    esc = html.escape
    mesurage = infos.get("mesurage", "").strip() or "xx 2026"
    date_v   = esc(infos.get("date", ""))
    dossier  = esc(infos.get("dossier", ""))
    batiment = esc(infos.get("batiment", ""))
    adresse  = esc(infos.get("adresse", ""))
    proprio  = esc(infos.get("proprio", ""))
    cadastre = esc(infos.get("cadastre", ""))

    css = """
@page { size: A4 landscape; margin: 10mm 18mm; }
* { box-sizing: border-box; }
body { font-family: Arial, sans-serif; font-size: 10px; color: #222; margin: 0; }
.page { page-break-after: always; display: flex; flex-direction: column;
        min-height: 190mm; padding: 0 4mm; }
.page:last-child { page-break-after: auto; }
header { display: grid; grid-template-columns: 1fr 2fr 1fr; align-items: center;
         border-bottom: 2px solid #1a2a4a; padding-bottom: 6px; margin-bottom: 10px; gap: 8px; }
header .logos { display: flex; align-items: center; gap: 10px; }
header .logos img { max-height: 46px; max-width: 120px; object-fit: contain; }
header .infos { text-align: center; font-size: 9px; line-height: 1.35; }
header .infos .title { font-size: 12px; font-weight: bold; color: #1a2a4a;
                       text-transform: uppercase; margin-bottom: 3px; }
header .meta { text-align: right; font-size: 9px; line-height: 1.4; }
header .meta .row { margin-bottom: 2px; }
header .meta .k { color: #555; font-weight: bold; margin-right: 4px; }
table.surf { width: 100%; border-collapse: collapse; font-size: 9px;
             table-layout: fixed; }
table.surf th, table.surf td { border: 1px solid #888; padding: 3px 4px;
                                text-align: center; vertical-align: middle;
                                word-wrap: break-word; overflow-wrap: break-word; }
table.surf th { background: #DDEEFF; font-weight: bold; }
table.surf th.sc { background: #B0C4DE; }
table.surf th.sub { background: #E8F0E8; }
table.surf td.idx { text-align: left; }
table.surf tr.total-etage td { background: #DCE6F1; font-weight: bold; text-align: left; }
table.surf tr.total-etage td.num { text-align: center; }
table.surf tr.total td { background: #B8CCE4; font-weight: bold; text-align: left; }
table.surf tr.total td.num { text-align: center; }
.footer { display: flex; align-items: flex-start; gap: 12px; margin-top: 6px; }
.footer .notas { flex: 1; min-width: 0; }
.nota-intro { font-size: 9px; font-style: italic; color: #333;
              line-height: 1.5; border-left: 2px solid #b0c4de; padding-left: 6px; }
.nota { margin-top: 4px; font-size: 8px; color: #555; line-height: 1.4; }
.nota p { margin: 2px 0; }
.stamp { flex: 0 0 auto; display: flex; align-items: flex-start;
         justify-content: center; padding-left: 8px; }
.stamp img { max-height: 95px; max-width: 180px; object-fit: contain; }
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
  .page { background: #fff; width: 297mm; margin: 0 auto 14px auto;
          padding: 10mm 18mm; box-shadow: 0 1px 4px rgba(0,0,0,.12); }
  body:not([data-all]) .page { display: none; }
  body:not([data-all]) .page.active { display: flex; }
}
@media print { .tabs { display: none !important; } }
"""

    def render_header(sheet_title):
        logos_html = ""
        if logo_ge:
            logos_html += f'<img src="{logo_ge}" alt="">'
        if logo_rtaxes:
            logos_html += f'<img src="{logo_rtaxes}" alt="">'
        return f"""
<header>
  <div class="logos">{logos_html}</div>
  <div class="infos">
    <div class="title">{esc(sheet_title)}</div>
    <div><b>Bâtiment :</b> {batiment} &nbsp;·&nbsp; <b>Adresse :</b> {adresse}</div>
    <div><b>Propriétaire :</b> {proprio} &nbsp;·&nbsp; <b>Cadastre :</b> {cadastre}</div>
  </div>
  <div class="meta">
    <div class="row"><span class="k">Date :</span>{date_v}</div>
    <div class="row"><span class="k">Dossier :</span>{dossier}</div>
  </div>
</header>
"""

    def render_table(df, sc_spans):
        cols = list(df.columns)
        index_cols = [c for c in ("Etage", "Occupant") if c in cols]
        value_cols = [c for c in cols if c not in index_cols]
        n_cols = len(cols)

        # Largeur colonnes : index étroites, reste équitablement réparti
        col_widths = []
        for c in cols:
            if c == "Etage":
                col_widths.append("6%")
            elif c == "Occupant":
                col_widths.append("14%")
            elif c == "Total":
                col_widths.append("8%")
            else:
                col_widths.append("auto")
        colgroup = "<colgroup>" + "".join(f'<col style="width:{w}">' for w in col_widths) + "</colgroup>"

        # Ligne sur-catégories si présentes
        sc_row = ""
        if sc_spans:
            # Pour chaque colonne, déterminer si elle est couverte par un span
            covered = {}  # idx → (sc_name, span_length) pour la première col du span
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
                    i += 1  # déjà couvert
                else:
                    cells.append('<th class="sc"></th>')
                    i += 1
            sc_row = "<tr>" + "".join(cells) + "</tr>"

        # Ligne en-tête colonnes
        hdr_cells = []
        for c in cols:
            cs = str(c)
            if cs.startswith("[") and cs.endswith("]"):
                hdr_cells.append(f'<th class="sub">Sous-total</th>')
            else:
                hdr_cells.append(f"<th>{esc(cs)}</th>")
        hdr_row = "<tr>" + "".join(hdr_cells) + "</tr>"

        # Lignes de données
        body_rows = []
        for _, row in df.iterrows():
            occupant_val = str(row.get("Occupant", "")) if "Occupant" in cols else ""
            etage_val    = str(row.get("Etage", ""))
            is_total_etage = occupant_val == "— Total étage"
            is_grand      = etage_val == "TOTAL"

            if is_grand:
                label = "TOTAL"
                tds = []
                if "Occupant" in index_cols:
                    tds.append(f'<td colspan="{len(index_cols)}">{label}</td>')
                else:
                    tds.append(f"<td>{label}</td>")
                for c in value_cols:
                    tds.append(f'<td class="num">{_fmt_num(row[c])}</td>')
                body_rows.append(f'<tr class="total">{"".join(tds)}</tr>')
                continue

            if is_total_etage:
                label = f"{esc(etage_val)} — Total étage"
                tds = [f'<td colspan="{len(index_cols)}">{label}</td>']
                for c in value_cols:
                    tds.append(f'<td class="num">{_fmt_num(row[c])}</td>')
                body_rows.append(f'<tr class="total-etage">{"".join(tds)}</tr>')
                continue

            tds = []
            for c in index_cols:
                tds.append(f'<td class="idx">{esc(str(row[c]))}</td>')
            for c in value_cols:
                tds.append(f'<td class="num">{_fmt_num(row[c])}</td>')
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

    stamp_html = f'<div class="stamp"><img src="{tampon}" alt=""></div>' if tampon else ""

    pages = []
    tabs = []
    for idx, (type_su, content) in enumerate(output_tables.items()):
        sheet_title = glossary_surf.real_su_name.get(type_su, type_su)
        df       = content["data"]
        sc_spans = content.get("sc_spans", {})
        nota     = content.get("nota", [])

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
            + render_table(df, sc_spans)
            + footer_html
            + "</section>"
        )
        pages.append(page_html)

    tabs_html = (
        '<nav class="tabs">'
        + "".join(tabs)
        + '<a href="#" class="tab-link" data-target="__all__" '
          'style="margin-left:auto">Tout afficher</a>'
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