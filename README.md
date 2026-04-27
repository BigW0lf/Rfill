# Rfill — Analyse des Surfaces

Outil desktop (Python / Tkinter) pour importer des exports AutoCAD ou GeoGex,
catégoriser les surfaces par glisser-déposer, et générer des tableaux Excel
normés selon les standards français du mesurage immobilier.

---

## Fonctionnalités

- Import multi-fichiers AutoCAD `.xls` et GeoGex `.xlsx`
- Détection automatique des types de surface (SUB, SU, SHO, SPD, GLA, TAX…)
- Catégorisation par glisser-déposer avec glossaire persistant
- Sur-catégories paramétrables (Superficies utiles / communes / annexes)
- Export Excel multi-onglets : en-tête projet, pivot par étage/occupant/catégorie,
  sous-totaux par sur-catégorie, total général
- Journal horodaté des opérations

---

## Prérequis

- Python 3.9+
- Conda (recommandé) ou pip

---

## Installation

```bash
# Avec pip
pip install -r requirements.txt

# Avec conda
conda install pandas xlsxwriter openpyxl
```

---

## Utilisation

```bash
python Main.py
```

### Workflow

1. **Onglet 1 — Importer** : charger un ou plusieurs fichiers AutoCAD (`.xls`) ou GeoGex (`.xlsx`)
2. **Onglet 2 — Glossaire** : vérifier la catégorisation automatique, ajuster par glisser-déposer, sauvegarder
3. **Onglet 3 — Générer** : renseigner les infos projet et exporter le fichier `.xlsx`

---

## Structure du projet

```
rfill/
├── Main.py            # Interface Tkinter (SurfaceApp — 5 onglets)
├── process_surf.py    # Logique métier : extraction, pivot, export Excel
├── glossary_surf.py   # Référentiel des catégories (modifié à la sauvegarde)
├── requirements.txt
├── CLAUDE.md          # Instructions pour Claude Code
└── .vscode/
    └── settings.json  # Interpréteur conda
```

---

## Codes de surface supportés

| Code | Libellé |
|------|---------|
| SUB  | Surfaces Utiles Brutes |
| SU   | Surfaces Utiles |
| SUBL | Surfaces Utiles Brutes Locatives |
| SUN  | Surfaces Utiles Nettes |
| SHO  | Surfaces Hors Œuvre |
| SPD  | Surfaces De Plancher |
| GLA  | Surfaces Globales |
| TAX  | Synthèse Surfaces Réelles (Art. 324 M CGI) |
| TSB  | Tableau Surfaces Brutes |

---

## Format des fichiers d'entrée

**AutoCAD XLS** — les lignes de données ont un entier en colonne 0 (numéro de séquence).
Le type de surface est dans la colonne `Calque` (ex. `"SUB Contours"` → type `SUB`).

**GeoGex XLSX** — même structure attendue.

---

## Développement

```bash
# Lancer l'application
python Main.py
```

Pas de tests automatisés pour le moment. La validation se fait en chargeant un
fichier XLS réel et en vérifiant l'export Excel généré.

---

## Auteur

**Jules FAGUET** — © 2026 Rfill — Tous droits réservés
