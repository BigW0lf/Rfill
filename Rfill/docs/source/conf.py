"""Configuration Sphinx pour la documentation Rfill."""

import os
import sys

# Rendre les modules du projet importables par autodoc
sys.path.insert(0, os.path.abspath("../.."))

# ── Informations projet ────────────────────────────────────────────────────────
project   = "Rfill"
copyright = "2026, Jules Faguet"
author    = "Jules Faguet"
version   = "1.0"
release   = "1.0.0"
language  = "fr"

# ── Extensions ────────────────────────────────────────────────────────────────
extensions = [
    "sphinx.ext.autodoc",        # génération depuis les docstrings
    "sphinx.ext.napoleon",       # support style Google / NumPy
    "sphinx.ext.viewcode",       # liens [source] vers le code
    "sphinx.ext.intersphinx",    # liens croisés vers pandas, Python std
    "sphinx.ext.autosummary",    # tableaux de résumé automatiques
    "sphinx_copybutton",         # bouton copier sur les blocs de code
]

# ── Napoleon (style Google) ───────────────────────────────────────────────────
napoleon_google_docstring    = True
napoleon_numpy_docstring     = False
napoleon_include_private_with_doc = True
napoleon_use_rtype           = True

# ── Autodoc ───────────────────────────────────────────────────────────────────
autodoc_default_options = {
    "members":          True,
    "undoc-members":    False,
    "private-members":  True,
    "show-inheritance": True,
    "member-order":     "bysource",
}
autodoc_typehints = "description"
add_module_names  = False

# ── Autosummary ───────────────────────────────────────────────────────────────
autosummary_generate = True

# ── Intersphinx ───────────────────────────────────────────────────────────────
intersphinx_mapping = {
    "python": ("https://docs.python.org/3", None),
    "pandas": ("https://pandas.pydata.org/docs", None),
}

# ── Thème Furo ────────────────────────────────────────────────────────────────
html_theme = "furo"
html_title = "Rfill — Documentation"

html_theme_options = {
    "sidebar_hide_name":    False,
    "navigation_with_keys": True,
    "light_css_variables": {
        "color-brand-primary":    "#1a2a4a",
        "color-brand-content":    "#2a4a8a",
        "color-admonition-background": "#f0f4ff",
    },
    "dark_css_variables": {
        "color-brand-primary":  "#7aacf0",
        "color-brand-content":  "#90bbf5",
    },
}

# ── Fichiers statiques ────────────────────────────────────────────────────────
html_static_path = ["_static"]
html_css_files   = ["custom.css"]

# ── Options diverses ──────────────────────────────────────────────────────────
exclude_patterns  = ["_build", "Thumbs.db", ".DS_Store"]
templates_path    = ["_templates"]
pygments_style    = "friendly"
