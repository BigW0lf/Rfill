Main — Interface graphique
==========================

Ce module contient la classe :class:`~Main.SurfaceApp`, point d'entrée unique
de l'application.  Il n'expose aucune API publique en dehors de cette classe.

.. contents:: Contenu
   :local:
   :depth: 2

Classe SurfaceApp
-----------------

.. autoclass:: Main.SurfaceApp
   :no-members:

Onglet Importer
~~~~~~~~~~~~~~~

.. automethod:: Main.SurfaceApp.build_import_tab
.. automethod:: Main.SurfaceApp.load_files
.. automethod:: Main.SurfaceApp.clear_data

Onglet Glossaire
~~~~~~~~~~~~~~~~

.. automethod:: Main.SurfaceApp.build_glossaire_tab
.. automethod:: Main.SurfaceApp.update_glossaire_tab
.. automethod:: Main.SurfaceApp.load_glossaire_board
.. automethod:: Main.SurfaceApp._make_aff_widget

Drag & drop — Affectations
""""""""""""""""""""""""""

.. automethod:: Main.SurfaceApp.start_drag
.. automethod:: Main.SurfaceApp.do_drag
.. automethod:: Main.SurfaceApp.stop_drag

Drag & drop — Colonnes
"""""""""""""""""""""""

.. automethod:: Main.SurfaceApp._cat_drag_start
.. automethod:: Main.SurfaceApp._cat_drag_do
.. automethod:: Main.SurfaceApp._cat_drag_stop

Gestion des catégories
""""""""""""""""""""""

.. automethod:: Main.SurfaceApp.add_category
.. automethod:: Main.SurfaceApp.rename_category
.. automethod:: Main.SurfaceApp.delete_category
.. automethod:: Main.SurfaceApp.save_glossaire
.. automethod:: Main.SurfaceApp.show_supercat_menu
.. automethod:: Main.SurfaceApp.move_cat_to_supercat
.. automethod:: Main.SurfaceApp.rename_supercat

Onglet Générer
~~~~~~~~~~~~~~

.. automethod:: Main.SurfaceApp.build_generer_tab
.. automethod:: Main.SurfaceApp.generate_final_tables

Onglets informatifs
~~~~~~~~~~~~~~~~~~~

.. automethod:: Main.SurfaceApp.build_aide_tab
.. automethod:: Main.SurfaceApp.build_apropos_tab
