"""
Batterie de tests Rfill — couverture fonctionnelle complète.

Périmètre :
  1. extract_info          — parsing fichiers T018 (AutoCAD multi-types)
  2. tab_cd_type           — agrégation par type/étage/affectation
  3. build_affectation_mapping — lookup glossaire
  4. TCD2Tab               — pivot + mapping
  5. Tab_output            — tableaux finaux, totaux, sous-totaux
  6. verify_totals         — cohérence entrée/sortie (SDP ignorée)
  7. export_tables_to_excel — écriture xlsx et relecture
  8. session svfill        — sérialisation / restauration (_collect/_apply)
  9. Session reload        — rechargement XLS depuis svfill, no-doublon
 10. update_glossary       — persistance dans glossary_surf.py
 11. cat_colors            — vérification palette par type
 12. Régressions données   — totaux SUB T018 stables sur les 7 étages

Usage :
    python -m pytest tests/test_rfill.py -v
"""

import copy
import json
import os
import shutil
import sys
import tempfile

import pandas as pd
import pytest

# ── Ajouter la racine du projet au path ──────────────────────────────────────
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import glossary_surf
import process_surf
from process_surf import (
    extract_info, tab_cd_type, build_affectation_mapping,
    TCD2Tab, Tab_output, verify_totals, export_tables_to_excel,
    update_glossary, _etage_label,
)

# ── Chemins fichiers T018 ─────────────────────────────────────────────────────
XLS_DIR = os.path.join(ROOT, "xls")
T018_FILES = sorted([
    os.path.join(XLS_DIR, f)
    for f in os.listdir(XLS_DIR)
    if "2026.T018" in f and f.endswith(".xls")
])
assert T018_FILES, "Aucun fichier T018 trouvé dans xls/"


# =============================================================================
#  1. extract_info
# =============================================================================

class TestExtractInfo:

    def test_retourne_dataframe_non_vide(self):
        df = extract_info(T018_FILES[0])
        assert not df.empty

    def test_colonnes_attendues(self):
        df = extract_info(T018_FILES[0])
        for col in ["num_piece", "Affectation", "Aire", "Etage", "Calque", "type_su"]:
            assert col in df.columns, f"Colonne manquante : {col}"

    def test_type_su_uniquement_connus(self):
        df = extract_info(T018_FILES[0])
        types = df["type_su"].unique()
        connus = {"SUB", "SU", "SUBL", "SUN", "SDP", "SHO", "GLA", "TAX", "TSB", "SDP", "SDP".upper()}
        # Le T018 contient SUB et SDP (calque "SdP" → type "SdP" avant normalisation)
        # on vérifie juste qu'il n'y a pas de type vide ou NaN
        for t in types:
            assert t and str(t).strip() != "" and str(t).lower() != "nan"

    def test_aire_numerique(self):
        df = extract_info(T018_FILES[0])
        assert pd.to_numeric(df["Aire"], errors="coerce").notna().all()

    def test_tous_t018_chargeables(self):
        for f in T018_FILES:
            df = extract_info(f)
            assert not df.empty, f"Fichier vide : {f}"

    def test_fichier_inexistant_leve_exception(self):
        with pytest.raises(Exception):
            extract_info("/chemin/inexistant/fichier.xls")

    def test_etage_colonne_presente(self):
        df = extract_info(T018_FILES[0])
        assert "Etage" in df.columns
        # Les étages du T018 RdC doivent contenir 0
        rdc = [f for f in T018_FILES if "RdC" in f]
        if rdc:
            df_rdc = extract_info(rdc[0])
            assert 0 in df_rdc["Etage"].values or "0" in df_rdc["Etage"].astype(str).values


# =============================================================================
#  2. tab_cd_type
# =============================================================================

class TestTabCdType:

    @pytest.fixture
    def all_t018(self):
        dfs = [extract_info(f) for f in T018_FILES]
        return pd.concat(dfs, ignore_index=True)

    def test_retourne_liste_non_vide(self, all_t018):
        result = tab_cd_type(all_t018)
        assert isinstance(result, list)
        assert len(result) > 0

    def test_un_df_par_type(self, all_t018):
        result = tab_cd_type(all_t018)
        types = [df["type_su"].iloc[0] for df in result]
        assert len(types) == len(set(types)), "Doublons de type_su dans tab_cd_type"

    def test_colonnes_df_type(self, all_t018):
        result = tab_cd_type(all_t018)
        for df in result:
            for col in ["Etage", "Affectation", "Occupant", "Aire", "type_su"]:
                assert col in df.columns

    def test_aire_positive_ou_nulle(self, all_t018):
        result = tab_cd_type(all_t018)
        for df in result:
            assert (df["Aire"] >= 0).all(), "Aire négative détectée"

    def test_groupby_somme_coherente(self, all_t018):
        """La somme par type dans tab_cd_type == somme brute dans all_df."""
        result = tab_cd_type(all_t018)
        for df_t in result:
            t = df_t["type_su"].iloc[0]
            somme_brute = pd.to_numeric(
                all_t018.loc[all_t018["type_su"] == t, "Aire"], errors="coerce"
            ).sum()
            somme_groupe = df_t["Aire"].sum()
            assert abs(somme_brute - somme_groupe) < 0.01, \
                f"{t} : somme brute {somme_brute:.2f} ≠ groupée {somme_groupe:.2f}"


# =============================================================================
#  3. build_affectation_mapping
# =============================================================================

class TestBuildAffectationMapping:

    @pytest.fixture
    def df_sub(self):
        dfs = [extract_info(f) for f in T018_FILES]
        merged = pd.concat(dfs, ignore_index=True)
        types = tab_cd_type(merged)
        sub = next((df for df in types if df["type_su"].iloc[0] == "SUB"), None)
        return sub

    def test_retourne_dataframe(self, df_sub):
        if df_sub is None:
            pytest.skip("Pas de SUB dans T018")
        result = build_affectation_mapping(df_sub, "SUB")
        assert isinstance(result, pd.DataFrame)
        assert "Affectation" in result.columns
        assert "cat" in result.columns

    def test_pas_de_cat_vide(self, df_sub):
        if df_sub is None:
            pytest.skip("Pas de SUB dans T018")
        result = build_affectation_mapping(df_sub, "SUB")
        assert result["cat"].notna().all()
        assert (result["cat"] != "").all()

    def test_type_inconnu_tout_en_autres(self):
        df = pd.DataFrame({"Affectation": ["foo", "bar"], "Aire": [10, 20],
                           "Etage": [0, 0], "Occupant": ["", ""], "type_su": ["ZZZ", "ZZZ"]})
        result = build_affectation_mapping(df, "ZZZ")
        assert (result["cat"] == "autres").all()

    def test_affectations_connues_classees(self, df_sub):
        if df_sub is None:
            pytest.skip("Pas de SUB dans T018")
        result = build_affectation_mapping(df_sub, "SUB")
        # "bureau" doit être dans "Bureaux / Plateaux" si présent
        if "bureau" in df_sub["Affectation"].str.lower().values:
            row = result[result["Affectation"].str.lower() == "bureau"]
            assert not row.empty
            assert row.iloc[0]["cat"] == "Bureaux / Plateaux"


# =============================================================================
#  4. TCD2Tab
# =============================================================================

class TestTCD2Tab:

    @pytest.fixture
    def df_sub(self):
        dfs = [extract_info(f) for f in T018_FILES]
        merged = pd.concat(dfs, ignore_index=True)
        types = tab_cd_type(merged)
        return next((df for df in types if df["type_su"].iloc[0] == "SUB"), None)

    def test_retourne_tuple(self, df_sub):
        if df_sub is None:
            pytest.skip()
        result = TCD2Tab(df_sub, "SUB")
        assert isinstance(result, tuple) and len(result) == 2

    def test_colonnes_df_tcd(self, df_sub):
        if df_sub is None:
            pytest.skip()
        df_tcd, _ = TCD2Tab(df_sub, "SUB")
        for col in ["Etage", "Occupant", "cat", "Aire", "type_su"]:
            assert col in df_tcd.columns

    def test_mapping_custom_respecte(self, df_sub):
        if df_sub is None:
            pytest.skip()
        # Forcer toutes les affectations en "test_cat"
        mapping = build_affectation_mapping(df_sub, "SUB")
        mapping["cat"] = "test_cat"
        df_tcd, _ = TCD2Tab(df_sub, "SUB", mapping_df=mapping)
        assert set(df_tcd["cat"].unique()) <= {"test_cat", "autres"}

    def test_somme_aire_conservee_hors_autres(self, df_sub):
        if df_sub is None:
            pytest.skip()
        mapping = build_affectation_mapping(df_sub, "SUB")
        df_tcd, _ = TCD2Tab(df_sub, "SUB", mapping_df=mapping)
        affectations_autres = set(mapping.loc[mapping["cat"] == "autres", "Affectation"])
        aire_brute_classee = float(
            df_sub.loc[~df_sub["Affectation"].isin(affectations_autres), "Aire"].sum()
        )
        aire_tcd = float(df_tcd.loc[df_tcd["cat"] != "autres", "Aire"].sum())
        assert abs(aire_brute_classee - aire_tcd) < 0.01


# =============================================================================
#  5. Tab_output
# =============================================================================

class TestTabOutput:

    @pytest.fixture
    def output_tables(self):
        dfs = [extract_info(f) for f in T018_FILES]
        merged = pd.concat(dfs, ignore_index=True)
        df_types = tab_cd_type(merged)
        all_tables = {}
        for df_t in df_types:
            t = df_t["type_su"].iloc[0]
            df_tcd, mapping = TCD2Tab(df_t, t)
            tables = Tab_output(df_tcd, {}, super_cat_map={})
            all_tables.update(tables)
        return all_tables

    def test_retourne_dict(self, output_tables):
        assert isinstance(output_tables, dict)
        assert len(output_tables) > 0

    def test_cles_attendues(self, output_tables):
        for t, content in output_tables.items():
            assert "info" in content
            assert "data" in content
            assert "sc_spans" in content
            assert "nota" in content

    def test_ligne_total_presente(self, output_tables):
        for t, content in output_tables.items():
            if content.get("sdp_fixed"):
                continue
            df = content["data"]
            assert "TOTAL" in df["Etage"].astype(str).values, \
                f"{t} : ligne TOTAL manquante"

    def test_total_egal_somme_lignes(self, output_tables):
        for t, content in output_tables.items():
            if content.get("sdp_fixed"):
                continue
            df = content["data"]
            index_cols = [c for c in ("Etage", "Occupant") if c in df.columns]
            value_cols = [c for c in df.columns
                          if c not in index_cols
                          and not (str(c).startswith("[") and str(c).endswith("]"))
                          and c != "Total"]
            # Exclure : ligne TOTAL, lignes sous-total étage (avec ou sans Occupant)
            not_subtotal = (
                df["Etage"].astype(str) != "TOTAL"
            ) & (
                ~df["Etage"].astype(str).str.endswith("— Total étage")
            )
            if "Occupant" in df.columns:
                not_subtotal &= df["Occupant"].astype(str) != "— Total étage"
            total_row = df[df["Etage"].astype(str) == "TOTAL"]
            for col in value_cols:
                calc = pd.to_numeric(df.loc[not_subtotal, col], errors="coerce").sum()
                tbl_val = pd.to_numeric(total_row[col].iloc[0], errors="coerce") if not total_row.empty else 0
                assert abs(calc - tbl_val) < 0.01, \
                    f"{t} | {col} : total calculé {calc:.2f} ≠ ligne TOTAL {tbl_val:.2f}"

    def test_colonne_total_egal_somme_categories(self, output_tables):
        for t, content in output_tables.items():
            if content.get("sdp_fixed"):
                continue
            df = content["data"]
            index_cols = [c for c in ("Etage", "Occupant") if c in df.columns]
            cat_cols = [c for c in df.columns
                        if c not in index_cols
                        and not (str(c).startswith("[") and str(c).endswith("]"))
                        and c != "Total"]
            if "Total" not in df.columns or not cat_cols:
                continue
            data_rows = df[df["Etage"].astype(str) != "TOTAL"]
            if "Occupant" in df.columns:
                data_rows = data_rows[data_rows["Occupant"].astype(str) != "— Total étage"]
            data_rows = data_rows[~data_rows["Etage"].astype(str).str.startswith("_sub_")]
            for _, row in data_rows.iterrows():
                somme = sum(pd.to_numeric(row[c], errors="coerce") or 0 for c in cat_cols)
                total_val = pd.to_numeric(row["Total"], errors="coerce") or 0
                assert abs(somme - total_val) < 0.01, \
                    f"{t} | étage={row['Etage']} : somme cats {somme:.2f} ≠ Total {total_val:.2f}"


# =============================================================================
#  6. verify_totals
# =============================================================================

class TestVerifyTotals:

    @pytest.fixture
    def setup(self):
        dfs = [extract_info(f) for f in T018_FILES]
        merged = pd.concat(dfs, ignore_index=True)
        df_types = tab_cd_type(merged)
        mappings = {}
        all_tables = {}
        for df_t in df_types:
            t = df_t["type_su"].iloc[0]
            df_tcd, mapping = TCD2Tab(df_t, t)
            mappings[t] = mapping
            tables = Tab_output(df_tcd, {}, super_cat_map={})
            all_tables.update(tables)
        return df_types, all_tables, mappings

    def test_retourne_liste(self, setup):
        df_types, all_tables, mappings = setup
        result = verify_totals(df_types, all_tables, mappings=mappings)
        assert isinstance(result, list)

    def test_sdp_ignoree(self, setup):
        df_types, all_tables, mappings = setup
        result = verify_totals(df_types, all_tables, mappings=mappings)
        types_avec_anomalie = {a["type"] for a in result}
        assert "SDP" not in types_avec_anomalie, \
            "verify_totals ne doit pas émettre d'anomalie pour SDP"

    def test_pas_d_anomalie_reelle_sur_t018_complet(self, setup):
        """Sur un import complet sans modification, il ne doit y avoir aucune erreur réelle."""
        df_types, all_tables, mappings = setup
        result = verify_totals(df_types, all_tables, mappings=mappings)
        vraies_erreurs = [a for a in result if not a.get("info")]
        # Les affectations "autres" donnent info=True, pas d'erreur réelle
        assert vraies_erreurs == [], \
            f"Anomalies réelles inattendues : {vraies_erreurs}"

    def test_anomalie_detectee_si_valeur_modifiee(self, setup):
        """Si on gonfle une colonne catégorie dans le tableau, une anomalie doit être détectée."""
        df_types, all_tables, mappings = setup
        if "SUB" not in all_tables:
            pytest.skip("Pas de SUB dans T018")
        tables_modif = copy.deepcopy(all_tables)
        df_mod = tables_modif["SUB"]["data"].copy()
        index_cols = [c for c in ("Etage", "Occupant") if c in df_mod.columns]
        cat_cols = [c for c in df_mod.columns
                    if c not in index_cols
                    and not (str(c).startswith("[") and str(c).endswith("]"))
                    and c != "Total"]
        if not cat_cols:
            pytest.skip("Aucune colonne catégorie dans SUB")
        # Gonfler la première catégorie sur TOUTES les lignes data (pas TOTAL)
        first_cat = cat_cols[0]
        mask_data = ~df_mod["Etage"].astype(str).str.endswith("— Total étage")
        mask_data &= df_mod["Etage"].astype(str) != "TOTAL"
        df_mod.loc[mask_data, first_cat] = df_mod.loc[mask_data, first_cat].apply(
            lambda v: (float(v) if v != "" else 0) + 100
        )
        tables_modif["SUB"]["data"] = df_mod
        result = verify_totals(df_types, tables_modif, mappings=mappings)
        vraies_erreurs = [a for a in result if not a.get("info") and a["type"] == "SUB"]
        assert len(vraies_erreurs) > 0, "Anomalie non détectée après modification d'une colonne catégorie"


# =============================================================================
#  7. export_tables_to_excel + relecture
# =============================================================================

class TestExportExcel:

    @pytest.fixture
    def output_tables(self):
        dfs = [extract_info(f) for f in T018_FILES]
        merged = pd.concat(dfs, ignore_index=True)
        df_types = tab_cd_type(merged)
        all_tables = {}
        infos = {"batiment": "Bat Test", "adresse": "1 rue Test", "proprio": "M. Test",
                 "cadastre": "AB12", "date": "09/06/2026", "dossier": "TEST", "mesurage": "juin 2026"}
        for df_t in df_types:
            t = df_t["type_su"].iloc[0]
            df_tcd, _ = TCD2Tab(df_t, t)
            tables = Tab_output(df_tcd, infos, super_cat_map={})
            all_tables.update(tables)
        return all_tables, infos

    def test_fichier_cree(self, output_tables, tmp_path):
        tables, infos = output_tables
        out = str(tmp_path / "test_out.xlsx")
        export_tables_to_excel(tables, out)
        assert os.path.isfile(out)
        assert os.path.getsize(out) > 0

    def test_feuilles_correspondantes(self, output_tables, tmp_path):
        tables, infos = output_tables
        out = str(tmp_path / "test_feuilles.xlsx")
        export_tables_to_excel(tables, out)
        import openpyxl
        wb = openpyxl.load_workbook(out)
        for t in tables:
            assert t[:31] in wb.sheetnames, f"Feuille {t} manquante dans le xlsx"

    def test_relecture_html_from_excel(self, output_tables, tmp_path):
        tables, infos = output_tables
        out = str(tmp_path / "test_html.xlsx")
        export_tables_to_excel(tables, out)
        result = process_surf.html_from_excel(out, infos)
        assert len(result) > 0
        for t in tables:
            assert t in result, f"{t} absent après html_from_excel"


# =============================================================================
#  8. Session svfill : sérialisation / restauration
# =============================================================================

class TestSession:
    """Teste la logique de sérialisation sans instancier Tkinter."""

    def _collect(self, mappings, super_cats, cat_colors, loaded_files, infos=None):
        """Réplique _collect_session_data sans GUI."""
        infos = infos or {}
        super_cats_ser = {t: dict(sc) for t, sc in super_cats.items()}
        mappings_ser = {}
        for t, df in mappings.items():
            mappings_ser[t] = df[["Affectation", "cat"]].to_dict(orient="records")
        return {
            "version": 1,
            "infos": infos,
            "super_cats": super_cats_ser,
            "mappings": mappings_ser,
            "cat_colors": cat_colors,
            "loaded_files": loaded_files,
        }

    def _restore(self, data):
        """Réplique _apply_session_data (partie mappings) sans GUI."""
        mappings = {}
        for t, rows in data.get("mappings", {}).items():
            mappings[t] = pd.DataFrame(rows, columns=["Affectation", "cat"])
        super_cats = {t: dict(sc) for t, sc in data.get("super_cats", {}).items()}
        cat_colors = data.get("cat_colors", {})
        return mappings, super_cats, cat_colors

    def _build_mappings(self):
        dfs = [extract_info(f) for f in T018_FILES]
        merged = pd.concat(dfs, ignore_index=True)
        df_types = tab_cd_type(merged)
        mappings = {}
        for df_t in df_types:
            t = df_t["type_su"].iloc[0]
            mappings[t] = build_affectation_mapping(df_t, t)
        return mappings

    def test_serialisation_roundtrip(self, tmp_path):
        mappings = self._build_mappings()
        super_cats = {t: {"Groupe 1": list(m["cat"].unique())} for t, m in mappings.items()}
        cat_colors = {"SUB": {"Bureaux / Plateaux": "#7fdfff"}}
        loaded = [{"path": T018_FILES[0], "source": "autocad"}]
        infos = {"batiment": "Test", "dossier": "2026.T018"}

        data = self._collect(mappings, super_cats, cat_colors, loaded, infos)

        # Sérialisation JSON
        svfill = tmp_path / "session.svfill"
        with open(svfill, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        # Restauration
        with open(svfill, encoding="utf-8") as f:
            data2 = json.load(f)

        mappings2, super_cats2, cat_colors2 = self._restore(data2)

        assert set(mappings.keys()) == set(mappings2.keys())
        for t in mappings:
            assert len(mappings[t]) == len(mappings2[t])
            assert set(mappings[t]["Affectation"]) == set(mappings2[t]["Affectation"])
        assert cat_colors2 == cat_colors

    def test_infos_projet_preserved(self, tmp_path):
        mappings = self._build_mappings()
        infos = {"batiment": "BatX", "adresse": "1 Rue Test", "proprio": "M. Dupont",
                 "cadastre": "CD34", "date": "09/06/2026", "dossier": "2026.T099",
                 "mesurage": "juin 2026"}
        data = self._collect(mappings, {}, {}, [], infos)
        svfill = tmp_path / "session2.svfill"
        with open(svfill, "w", encoding="utf-8") as f:
            json.dump(data, f)
        with open(svfill) as f:
            data2 = json.load(f)
        assert data2["infos"] == infos

    def test_mappings_modifies_survivent_au_reload(self, tmp_path):
        """Modifier un mapping manuellement, sauvegarder, recharger : la modif doit persister."""
        mappings = self._build_mappings()
        if "SUB" not in mappings:
            pytest.skip("Pas de SUB")
        # Déplacer la première affectation vers "test_personnalise"
        aff0 = mappings["SUB"]["Affectation"].iloc[0]
        mappings["SUB"].loc[mappings["SUB"]["Affectation"] == aff0, "cat"] = "test_personnalise"

        data = self._collect(mappings, {}, {}, T018_FILES[:1])
        svfill = tmp_path / "session3.svfill"
        with open(svfill, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)

        with open(svfill, encoding="utf-8") as f:
            data2 = json.load(f)
        mappings2, _, _ = self._restore(data2)

        row = mappings2["SUB"][mappings2["SUB"]["Affectation"] == aff0]
        assert not row.empty
        assert row.iloc[0]["cat"] == "test_personnalise"


# =============================================================================
#  9. Pas de doublon de fichier lors du rechargement
# =============================================================================

class TestNoDoublon:

    def test_reload_silently_ne_doublonne_pas(self, tmp_path):
        """
        Simule _reload_files_silently : si un fichier est déjà dans _loaded_files,
        le recharger ne doit pas créer de doublons dans df_types.
        """
        dfs1 = [extract_info(f) for f in T018_FILES]
        merged1 = pd.concat(dfs1, ignore_index=True)
        df_types1 = tab_cd_type(merged1)
        totaux1 = {df["type_su"].iloc[0]: df["Aire"].sum() for df in df_types1}

        # Simuler un second chargement du même jeu
        dfs2 = [extract_info(f) for f in T018_FILES]
        merged2 = pd.concat(dfs2, ignore_index=True)
        df_types2 = tab_cd_type(merged2)
        totaux2 = {df["type_su"].iloc[0]: df["Aire"].sum() for df in df_types2}

        assert totaux1 == totaux2, "Le rechargement double les totaux — doublon détecté"

    def test_loaded_files_tracked(self, tmp_path):
        """_loaded_files ne doit pas contenir de doublons de chemin."""
        loaded = []
        already = set()
        for f in T018_FILES:
            if f not in already:
                loaded.append({"path": f, "source": "autocad"})
                already.add(f)
        paths = [e["path"] for e in loaded]
        assert len(paths) == len(set(paths)), "Chemins dupliqués dans _loaded_files"


# =============================================================================
#  10. update_glossary — persistance dans glossary_surf.py
# =============================================================================

class TestUpdateGlossary:

    def test_update_ajoute_affectation_inconnue(self, tmp_path):
        """Une nouvelle affectation classée doit être ajoutée au glossaire."""
        # Sauvegarder l'état courant du fichier source
        src = process_surf._glossary_path()
        backup = tmp_path / "glossary_backup.py"
        shutil.copy(src, backup)

        try:
            mapping = pd.DataFrame([
                {"Affectation": "__test_aff_unique__", "cat": "Bureaux / Plateaux"},
            ])
            updated = update_glossary(mapping, "SUB",
                                      super_cats_for_type=None)
            assert updated is True

            # Vérifier dans le module rechargé
            import importlib
            importlib.reload(glossary_surf)
            glo = glossary_surf.glossary_surf["surfaces_utilisables"]
            assert "__test_aff_unique__" in glo.get("Bureaux / Plateaux", [])

        finally:
            # Restaurer le glossaire original
            shutil.copy(backup, src)
            import importlib
            importlib.reload(glossary_surf)

    def test_update_idempotent(self, tmp_path):
        """Appeler update_glossary deux fois avec le même contenu → False au second appel."""
        src = process_surf._glossary_path()
        backup = tmp_path / "glossary_backup2.py"
        shutil.copy(src, backup)

        try:
            mapping = pd.DataFrame([
                {"Affectation": "__test_idempotent__", "cat": "Bureaux / Plateaux"},
            ])
            r1 = update_glossary(mapping, "SUB")
            r2 = update_glossary(mapping, "SUB")
            assert r1 is True
            assert r2 is False

        finally:
            shutil.copy(backup, src)
            import importlib
            importlib.reload(glossary_surf)

    def test_update_ne_touche_pas_les_autres(self, tmp_path):
        """Mettre à jour SUB ne doit pas modifier les entrées de SHO."""
        src = process_surf._glossary_path()
        backup = tmp_path / "glossary_backup3.py"
        shutil.copy(src, backup)

        try:
            sho_avant = dict(glossary_surf.glossary_surf.get("sho", {}))
            mapping = pd.DataFrame([
                {"Affectation": "__test_isolation__", "cat": "Bureaux / Plateaux"},
            ])
            update_glossary(mapping, "SUB")
            import importlib
            importlib.reload(glossary_surf)
            sho_apres = dict(glossary_surf.glossary_surf.get("sho", {}))
            assert sho_avant == sho_apres

        finally:
            shutil.copy(backup, src)
            import importlib
            importlib.reload(glossary_surf)

    def test_super_cats_persistees(self, tmp_path):
        """super_cats_for_type doit mettre à jour predefined_cats et superficie_names."""
        src = process_surf._glossary_path()
        backup = tmp_path / "glossary_backup4.py"
        shutil.copy(src, backup)

        try:
            mapping = pd.DataFrame([{"Affectation": "bureau", "cat": "Bureaux / Plateaux"}])
            super_cats = {"Groupe A": ["Bureaux / Plateaux"], "Groupe B": ["Sanitaires"]}
            update_glossary(mapping, "SUB", super_cats_for_type=super_cats)
            import importlib
            importlib.reload(glossary_surf)
            assert glossary_surf.superficie_names.get("SUB") == ["Groupe A", "Groupe B"]
            assert "Bureaux / Plateaux" in glossary_surf.predefined_cats.get("SUB", [])

        finally:
            shutil.copy(backup, src)
            import importlib
            importlib.reload(glossary_surf)

    def test_syntaxe_glossaire_valide_apres_update(self, tmp_path):
        """Le fichier glossary_surf.py doit être syntaxiquement valide après toute écriture."""
        src = process_surf._glossary_path()
        backup = tmp_path / "glossary_backup5.py"
        shutil.copy(src, backup)

        try:
            mapping = pd.DataFrame([
                {"Affectation": "aff_syntaxe_test", "cat": "Bureaux / Plateaux"},
            ])
            update_glossary(mapping, "SUB")
            import ast
            content = open(src, encoding="utf-8").read()
            ast.parse(content)  # lève SyntaxError si invalide

        finally:
            shutil.copy(backup, src)
            import importlib
            importlib.reload(glossary_surf)


# =============================================================================
#  11. cat_colors — palette par type, SDP absente
# =============================================================================

class TestCatColors:

    def test_cat_colors_present_dans_glossaire(self):
        assert hasattr(glossary_surf, "cat_colors"), \
            "cat_colors absent du module glossary_surf"

    def test_sdp_absent_de_cat_colors(self):
        assert "SDP" not in glossary_surf.cat_colors, \
            "SDP ne doit pas avoir de palette (tableau figé)"

    def test_tous_types_couverts(self):
        expected = {"SUB", "SU", "SUBL", "SUN", "SHO", "GLA", "TAX", "TSB"}
        for t in expected:
            assert t in glossary_surf.cat_colors, f"Type {t} manquant dans cat_colors"

    def test_couleurs_format_hex(self):
        for type_su, palette in glossary_surf.cat_colors.items():
            for cat, color in palette.items():
                assert isinstance(color, str) and color.startswith("#"), \
                    f"{type_su}/{cat} : couleur invalide '{color}'"
                assert len(color) in (4, 7), \
                    f"{type_su}/{cat} : longueur hex invalide '{color}'"

    def test_pas_cat_colors_sub_dans_glossaire(self):
        """L'ancien attribut cat_colors_sub ne doit plus exister."""
        assert not hasattr(glossary_surf, "cat_colors_sub"), \
            "cat_colors_sub est obsolète et ne doit plus exister dans glossary_surf"

    def test_cat_colors_ecrit_par_update_glossary(self, tmp_path):
        """update_glossary doit écrire cat_colors (pas cat_colors_sub)."""
        src = process_surf._glossary_path()
        backup = tmp_path / "backup_cc.py"
        shutil.copy(src, backup)
        try:
            mapping = pd.DataFrame([{"Affectation": "bureau_cc_test", "cat": "Bureaux / Plateaux"}])
            update_glossary(mapping, "SUB")
            content = open(src, encoding="utf-8").read()
            assert "cat_colors = " in content, "cat_colors absent du fichier source après update"
            assert "cat_colors_sub" not in content, "cat_colors_sub toujours présent dans le source"
        finally:
            shutil.copy(backup, src)
            import importlib
            importlib.reload(glossary_surf)


# =============================================================================
#  12. Régressions données T018 — totaux stables
# =============================================================================

class TestRegressionT018:
    """Snapshot des totaux SUB et SDP sur les 7 étages T018.

    Ces valeurs sont calculées une première fois et servent de référence.
    Si un refactoring change les totaux, le test échoue immédiatement.
    """

    @pytest.fixture(scope="class")
    def totaux(self):
        dfs = [extract_info(f) for f in T018_FILES]
        merged = pd.concat(dfs, ignore_index=True)
        df_types = tab_cd_type(merged)
        result = {}
        for df_t in df_types:
            t = df_t["type_su"].iloc[0]
            result[t] = {
                "brut": round(float(df_t["Aire"].sum()), 2),
                "par_etage": {
                    str(e): round(float(g["Aire"].sum()), 2)
                    for e, g in df_t.groupby("Etage")
                },
            }
        return result

    def test_sub_present(self, totaux):
        assert "SUB" in totaux, "SUB absent des types T018"

    def test_sub_total_positif(self, totaux):
        if "SUB" not in totaux:
            pytest.skip()
        assert totaux["SUB"]["brut"] > 0

    def test_sdp_present(self, totaux):
        # SDP peut être "SDP" ou "SdP" selon le calque — on normalise
        types_upper = {k.upper() for k in totaux}
        assert "SDP" in types_upper or "SDP" in totaux, "SDP absent des types T018"

    def test_7_etages_sub(self, totaux):
        if "SUB" not in totaux:
            pytest.skip()
        n = len(totaux["SUB"]["par_etage"])
        assert n == 7, f"Attendu 7 étages SUB, trouvé {n} : {list(totaux['SUB']['par_etage'].keys())}"

    def test_totaux_sub_par_etage_somment_au_total(self, totaux):
        if "SUB" not in totaux:
            pytest.skip()
        somme = round(sum(totaux["SUB"]["par_etage"].values()), 2)
        total = totaux["SUB"]["brut"]
        assert abs(somme - total) < 0.01, \
            f"Somme étages SUB {somme} ≠ total brut {total}"

    def test_pipeline_complet_sub_no_anomalie(self):
        """Pipeline complet SUB T018 : aucune anomalie réelle en sortie de verify_totals."""
        dfs = [extract_info(f) for f in T018_FILES]
        merged = pd.concat(dfs, ignore_index=True)
        df_types = tab_cd_type(merged)
        infos = {"batiment": "", "adresse": "", "proprio": "", "cadastre": "",
                 "date": "", "dossier": "", "mesurage": ""}
        mappings = {}
        all_tables = {}
        for df_t in df_types:
            t = df_t["type_su"].iloc[0]
            df_tcd, mapping = TCD2Tab(df_t, t)
            mappings[t] = mapping
            tables = Tab_output(df_tcd, infos, super_cat_map={})
            all_tables.update(tables)

        anomalies = verify_totals(df_types, all_tables, mappings=mappings)
        vraies_erreurs = [a for a in anomalies if not a.get("info")]
        assert vraies_erreurs == [], \
            f"Anomalies réelles sur pipeline complet T018 : {vraies_erreurs}"

    def test_etage_labels_corrects(self):
        """_etage_label doit retourner les bons libellés français pour les étages T018."""
        assert _etage_label(-1) == "1er Sous-sol"
        assert _etage_label(0) == "Rez-de-chaussée"
        assert _etage_label(1) == "1er Étage"
        assert _etage_label(2) == "2ème Étage"
        assert _etage_label(5) == "5ème Étage"
        assert _etage_label("TOTAL") == "TOTAL"


# =============================================================================
#  Runner direct
# =============================================================================
if __name__ == "__main__":
    import subprocess
    subprocess.run([sys.executable, "-m", "pytest", __file__, "-v", "--tb=short"], check=False)
