process\_surf — Traitement des données
======================================

Ce module contient toute la logique métier : lecture des fichiers Excel,
agrégation, catégorisation, génération du classeur Excel et du rapport HTML.

.. contents:: Contenu
   :local:
   :depth: 2

Fonctions utilitaires internes
------------------------------

.. autofunction:: process_surf._glossary_path

.. autofunction:: process_surf._reload_glossary

.. autofunction:: process_surf._img_dir

.. autofunction:: process_surf._img_to_base64

.. autofunction:: process_surf._fmt_num

Fonctions métier
----------------

Importation
~~~~~~~~~~~

.. autofunction:: process_surf.extract_info

.. autofunction:: process_surf.tab_cd_type

Catégorisation
~~~~~~~~~~~~~~

.. autofunction:: process_surf.build_affectation_mapping

.. autofunction:: process_surf.TCD2Tab

.. autofunction:: process_surf.update_glossary

Mise en forme et export
~~~~~~~~~~~~~~~~~~~~~~~

.. autofunction:: process_surf.Tab_output

.. autofunction:: process_surf.export_tables_to_excel

.. autofunction:: process_surf.export_tables_to_html
