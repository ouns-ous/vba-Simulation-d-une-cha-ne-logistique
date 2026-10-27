# Cartographie des boutons

Audit statique des gestionnaires d'événements dans `VBA_Source/F_Main.frm` et des routines appelées dans les modules VBA. Cette cartographie décrit le comportement défini dans les sources; elle ne certifie pas un nouveau clic manuel dans Excel.

## Navigation principale

| Bouton | Destination | Fonction |
|---|---|---|
| Accueil | Accueil | Retourne à la synthèse opérationnelle du formulaire. |
| nav1 | Produits | Catalogue et gestion des produits. |
| nav2 | Fournisseurs | Gestion des fournisseurs. |
| nav3 | Commandes | Gestion des commandes. |
| nav4 | Approvisionnements | Création et réception des approvisionnements. |
| nav5 | Stocks | Consultation et mouvements de stock. |
| nav6 | Paramètres | Paramètres de demande et sélection du produit à simuler. |
| nav7 | Simulation | Politiques s / Q, tirages et comparaison. |
| nav8 | Résultats | Résultats de la comparaison et décision. |
| nav9 | Export / Aide | Exports CSV, Power BI et guide. |
| Rapports | Rapports | Vues synthèse, évolution, coûts, comparaison et sensibilité. |

## Paramètres et simulation

| Bouton | Fonction |
|---|---|
| Enregistrer les paramètres / scénarios | Valide les champs et enregistre paramètres et politiques; marque les résultats périmés si les valeurs ont changé. |
| Recharger | Recharge les paramètres du classeur sans recalibrer la demande. |
| Réinitialiser | Remet les paramètres et politiques aux valeurs initiales, tout en conservant le produit actuellement sélectionné s'il existe; sinon P001 est choisi. |
| Calibrer | Enregistre les valeurs affichées puis recalcule les paramètres historiques du produit sélectionné. |
| Tirages normaux | Génère et affiche dix tirages de demande pour les paramètres actifs. |
| Comparer les scénarios | Compare les trois politiques sur la même matrice de demande, actualise les résultats et ouvre l'onglet Résultats. |
| Simuler les valeurs saisies s / Q | Exécute les répétitions pour la politique personnalisée et ouvre son rapport. |
| Résultats personnalisés | Réaffiche le rapport de la simulation personnalisée précédente. |
| Actualiser les résultats | Relance la comparaison si son état est périmé, puis recalcule l'affichage. |
| Synthèse / Évolution | Ouvre les rapports dédiés dans la page Rapports. |

Le sélecteur de produit charge les quatre références du catalogue. Changer de produit recalcule la calibration et invalide les résultats précédents. La comparaison reste mono-produit.

## Export et aide

| Bouton | Fonction |
|---|---|
| Exporter CSV | Exporte les tables et les résultats courants dans le dossier `exports`. |
| Power BI | Enregistre et valide les paramètres, actualise les exports, écrit le chemin local des exports et ouvre le projet PBIP. Il faut ensuite actualiser le modèle dans Power BI Desktop. |
| Voir le guide Excel | Masque le formulaire et active l'onglet Guide. Le texte du guide est réécrit à l'ouverture du formulaire. |
| Accueil / Fermer | Revient à Menu ou ferme le formulaire en restituant la fenêtre Excel. |

## Gestion des données

- Produits et fournisseurs : recherche, nouveau, sélection d'une ligne, enregistrement et suppression autorisée.
- Commandes : recherche, nouvelle commande, sélection, enregistrement et annulation. L'annulation suit les contrôles du module de données.
- Approvisionnements : recherche, nouveau, sélection, enregistrement et réception unique.
- Stocks : filtre par produit, nouveau mouvement, enregistrement et affichage de la table Excel.

## Contrôles de cette intervention

- Revue des gestionnaires `*_Click` de `F_Main.frm` et de leurs appels VBA.
- Vérification de la préservation du produit lors de Réinitialiser.
- Vérification de l'actualisation et de l'écriture de `Simulation_Contexte.csv` pendant Exporter CSV / Power BI.
- Le service de contrôle de fenêtre Excel n'a pas permis de reprendre une session interactive lors de cette intervention. Les clics réels des formulaires et le rendu visuel Power BI n'ont donc pas été revalidés. Les résultats antérieurs consignés dans `verification.json` et `AUDIT_FINAL.md` sont des essais précédents, pas une preuve de clics effectués aujourd'hui.

## Navigation du rapport Power BI

Le rapport conserve quatre pages : **Vue globale**, **Stock**, **Scénarios** et **Coûts**. Chaque page expose les quatre destinations dans sa navigation. Les entrées vers **Fournisseurs** et **Transactions**, leurs boutons, et les pages correspondantes ont été supprimés. Les tables historiques peuvent rester dans le modèle sémantique sans que ces pages soient présentes.

