# Suggestions de stock par produit

22 octobre 2026. Les suggestions sont calculées depuis les paramètres du produit choisi, après calibration de son historique. Les anciens couples fixes s/Q et S0=500 ont été retirés de la réinitialisation.

## Calcul

Avec mu et sigma journaliers, délai L, coût de commande Ccmd et coût de possession journalier Cpos :

- Quantile z de la normale pour l'objectif de service, borné entre 0,5 et 0,9999 pour accepter les objectifs de 0 et 1.
- Demande de protection D = mu × (L+1). Le jour supplémentaire tient compte de la revue quotidienne et de la commande après les ventes.
- Variabilité de protection V = sigma × racine(L+1).
- Lot de référence q = max(1, mu, racine(2 × mu × Ccmd / Cpos)). Si Ccmd ou Cpos est nul, utiliser max(1, mu, mu × L).
- Prudent : s = plafond(D + (z+0,75) × V), Q = plafond(1,35 × q).
- Flux tendu : s = plafond(D + max(0,z−0,75) × V), Q = plafond(max(mu,0,70 × q)).
- Équilibré : s = plafond(D + z × V), Q = plafond(q).
- S0 = plafond(s équilibré + Q équilibré / 2). Ce S0 est commun aux trois scénarios du produit. La simulation personnalisée reçoit initialement le couple équilibré.

Les suggestions sont des points de départ à comparer. Le quantile constitue une approximation pour construire le stock de sécurité. Il ne garantit pas le service en volume. La décision finale dépend des résultats simulés. Le moteur conserve le coût de rupture calibré pour évaluer les scénarios.

S0 est un stock initial théorique de simulation. Il ne modifie pas le stock physique de gestion dans T_Stock.

## Valeurs après calibration

Paramètres de référence : Cpos=0,03 USD/unité/jour, service 95 %. Les quatre calibrations donnent actuellement L=3 et Ccmd=12.

| Produit | Prudent s / Q | Flux tendu s / Q | Équilibré s / Q | S0 |
|---|---:|---:|---:|---:|
| P001 | 374 / 327 | 323 / 170 | 349 / 243 | 471 |
| P002 | 327 / 303 | 279 / 157 | 303 / 224 | 415 |
| P003 | 301 / 290 | 257 / 151 | 279 / 215 | 387 |
| P004 | 112 / 180 | 98 / 94 | 105 / 134 | 172 |

Le changement de produit et la recalibration régénèrent ces suggestions. Enregistrer ou recharger conserve les modifications manuelles. La réinitialisation conserve le produit sélectionné et recalcule ses valeurs. Changer S0 met à jour la colonne S0 des scénarios. Toute modification invalide les résultats précédents.

Les sources des modules concernés ont été injectées dans DataCo_Logistique.xlsm. Les exports CSV ont été recalculés pour P002, produit sélectionné avant intervention. Power BI charge ces nouveaux exports à sa prochaine actualisation.

## Vérifications

Comparaison native Excel réussie pour les quatre produits. Recalcul indépendant depuis les 87 094 commandes : moyenne, écart-type, douze couples s/Q, quatre S0 et 360 répétitions de simulation. Les coûts, services, stocks moyens et nombres de commandes concordent avec le moteur VBA.

Tests natifs du formulaire : sélection des quatre références, mise à jour des champs, conservation des saisies manuelles après enregistrement et rechargement, nouvelle suggestion après recalibration. Tests des événements de feuille : sélection du produit, invalidation des résultats et synchronisation de S0. Les coûts nuls, demande nulle, objectifs 0/1 et dépassement de capacité sont également vérifiés. Aucun module de test n'est conservé dans le classeur.

Contrôle du fichier sauvegardé : 19 feuilles conservées, six tables de gestion identiques à la sauvegarde avant intervention, aucune cellule en erreur, P002 sélectionné et résultats ACTUEL. Les tableaux de paramètres et de scénarios ont été contrôlés à partir d'un rendu PDF natif Excel.
