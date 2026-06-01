Introduction
============

**Rfill** est un outil desktop (Windows) destiné aux professionnels de
l'immobilier français.  Il importe les exports Excel d'AutoCAD (``.xls``) et
de GeoGex (``.xlsx``), permet de catégoriser les surfaces par glisser-déposer,
puis génère des tableaux de surfaces normés au format Excel (``.xlsx``) et un
rapport imprimable A4 paysage (``.html``).

Normes gérées
-------------

.. list-table::
   :header-rows: 1
   :widths: 10 90

   * - Code
     - Désignation
   * - **SUB**
     - Surfaces Utiles Brutes
   * - **SU**
     - Surfaces Utiles
   * - **SUBL**
     - Surfaces Utiles Brutes Locatives
   * - **SUN**
     - Surfaces Utiles Nettes
   * - **SHO**
     - Surfaces Hors Œuvre
   * - **SDP**
     - Surfaces De Plancher
   * - **GLA**
     - Surfaces Globales
   * - **TAX**
     - Synthèse des Surfaces Réelles (Art. 324 M/Z CGI)
   * - **TSB**
     - Tableau Surfaces Brutes

Architecture
------------

Le projet est organisé en trois modules :

.. code-block:: text

    Main.py            Interface Tkinter (SurfaceApp) — 5 onglets
    process_surf.py    Traitement des données, exports Excel et HTML
    glossary_surf.py   Données de référence statiques (glossaires, libellés, notas)

Flux de données
---------------

.. code-block:: text

    Fichiers .xls/.xlsx
          │
          ▼
    extract_info()          Lecture + normalisation
          │
          ▼
    tab_cd_type()           Agrégation par type / étage / occupant
          │
          ▼
    build_affectation_mapping()   Lookup automatique dans le glossaire
          │
          ▼
    [Interface Tkinter]     Ajustements manuels par glisser-déposer
          │
          ▼
    TCD2Tab()               Fusion mapping ↔ données
          │
          ▼
    Tab_output()            Pivot + totaux + spans sur-catégories
          │
          ├─► export_tables_to_excel()   Classeur .xlsx multi-feuilles
          └─► export_tables_to_html()    Rapport A4 paysage auto-contenu

Prérequis
---------

.. code-block:: bash

    pip install pandas xlsxwriter openpyxl

Lancement
---------

.. code-block:: bash

    python Main.py
