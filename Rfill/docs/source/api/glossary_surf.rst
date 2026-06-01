glossary\_surf — Données de référence
======================================

Module de données pures.  Toutes les variables sont des constantes chargées
à l'import.  Le fichier est **modifiable par l'utilisateur** : le bouton
*Sauvegarder* de l'onglet Glossaire y ajoute les nouvelles affectations
via :func:`process_surf.update_glossary`.

.. contents:: Contenu
   :local:

.. automodule:: glossary_surf
   :no-members:

Glossaires de catégorisation
-----------------------------

.. data:: glossary_surf.glossary_surf

   Glossaires de mots-clés par famille de surface.

   Structure ::

       {
           "surfaces_utilisables": {
               "bureaux / plateaux": ["bureau", "open space", …],
               …
           },
           "sho": { … },
           "spd": { … },
       }

   Les familles disponibles sont ``surfaces_utilisables``, ``sho`` et ``spd``.

Correspondances codes → ressources
------------------------------------

.. data:: glossary_surf.denom_surf

   Mappe chaque code de surface vers la clé de glossaire applicable.

   Exemple : ``{"SUB": "surfaces_utilisables", "SHO": "sho", …}``.

.. data:: glossary_surf.real_su_name

   Libellés complets pour les en-têtes des tableaux.

   Exemple : ``{"SU": "SURFACES UTILES (SU*)", …}``.

.. data:: glossary_surf.superficie_names

   Noms des super-catégories par défaut, par code de surface.

   Exemple : ``{"SU": ["Superficies utiles", "Superficies annexes"], …}``.

.. data:: glossary_surf.predefined_cats

   Catégories pré-saisies pour certains types (SHO, SUBL, SUN, TAX, SDP).
   Affichées dans le board Glossaire même si aucune donnée ne les contient.

Notes réglementaires
---------------------

.. data:: glossary_surf.nota_surf

   Notes légales et définitions affichées en bas des tableaux Excel et HTML,
   par code de surface (SDP, SU, TAX, SUB, SUN, SUBL, SHO).
