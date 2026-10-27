# DataCo Global

Ouvrir **DataCo_Logistique.xlsm** dans Excel Desktop et autoriser ses macros. Le formulaire de gestion s'ouvre automatiquement.

Le premier onglet Excel, **Menu**, contient le bouton **Ouvrir le formulaire** et les raccourcis de gestion. Fermer le formulaire avec **Fermer** ou sa croix rend le classeur visible sur Menu. Les feuilles **T_Fournisseurs, T_Produits, T_Commandes, T_Stock** et **T_Approvisionnements** sont accessibles directement après Menu. Les tables sont protégées contre les modifications accidentelles ; utiliser les formulaires pour les opérations de gestion.

## Simulation et rapports

1. Dans **Paramètres**, choisir l'un des quatre produits. Le changement de sélection recalcule sa demande, ses paramètres historiques, son **S0 de simulation** et les **trois suggestions s/Q**. La simulation porte sur un seul produit à la fois.
2. Dans **Simulation**, consulter ou ajuster les seuils et quantités propres au produit, puis cliquer **Comparer et afficher les 3 scénarios**. **Recalibrer et suggérer s / Q / S0** régénère les suggestions depuis l'historique. Recharger ou enregistrer conserve les valeurs saisies. Un changement de produit régénère les suggestions pour la nouvelle référence.
3. **Résultats** présente les moyennes, les intervalles de confiance du coût et la décision. **Actualiser** recalcule les résultats périmés.
4. Dans **Rapports**, choisir une des six vues : synthèse, journal quotidien du stock, disponibilité et coûts, comparaisons appariées, sensibilité Oui/Non, simulation personnalisée.
5. Pour les valeurs personnalisées de s et Q, cliquer **Simuler les valeurs saisies s / Q**. Le rapport affiche les répétitions correspondantes. Après un changement de produit ou de paramètres, relancer la simulation ou la comparaison avant d'interpréter les résultats.

Les tableaux utilisent toute la largeur disponible, avec en-têtes alignés et défilement vertical des listes longues. Le formulaire garde ses boutons Accueil et Fermer en bas. Une modification des paramètres ou une opération de gestion signale que les résultats de comparaison sont périmés.

## Power BI

Le rapport utilise désormais le thème bleu et blanc inspiré de la référence fournie : fond bleu très clair, cartes blanches arrondies, titres bleu marine, quatre KPI en tête, panneau de contrôle à droite et navigation en bas. Les quatre pages et les calculs existants sont conservés. La vue globale inclut un tableau de synthèse des scénarios. Le panneau affiche le produit réellement simulé ; les simulations se relancent depuis Excel.

Ouvrir **PowerBI/DataCo.pbip** pour consulter le rapport. Le thème importable séparément est **PowerBI/DataCo_Theme_BleuBlanc.json**. Une sauvegarde du rapport avant modification est conservée dans **PowerBI/sauvegardes/DataCo_avant_theme_20261022_092110.zip**. Le contrôle **PowerBI/verification_theme.json** confirme la validation des schémas Microsoft, des formats de visuels, des références aux champs et des limites du canevas, sans chevauchement des éléments de contenu. La prévisualisation native dans Power BI Desktop n'a pas pu être vérifiée : l'outil de contrôle du bureau n'a pas démarré.

Le bouton **Analyse décisionnelle - Power BI** du Menu, ou le bouton Power BI de **Export / Aide**, actualise les CSV, configure le chemin local des exports puis ouvre **PowerBI/DataCo.pbip**. L'export inclut `Simulation_Contexte.csv` : le code et le nom du produit simulé sont affichés sur chacune des quatre pages. Dans Power BI Desktop, cliquer **Actualiser** pour charger les derniers exports. Les résultats suivent la dernière comparaison à jour pour le produit choisi.

Le rapport Power BI contient quatre pages : **Vue globale**, **Stock**, **Scénarios** et **Coûts**. Les pages **Fournisseurs** et **Transactions** ont été retirées du rapport. Chaque page garde une barre de navigation vers les quatre pages restantes. En mode édition de Power BI Desktop, utiliser **Ctrl + clic** ; en mode Lecture, un clic suffit. Cette distinction est documentée par [Microsoft](https://learn.microsoft.com/en-us/power-bi/create-reports/button-navigators).

Après un déplacement du dossier, le bouton Excel réécrit automatiquement le paramètre TMDL `FolderExports`. Pour ouvrir directement le projet Power BI après déplacement, fermer Power BI puis exécuter **Configurer_PowerBI.ps1** depuis ce dossier.

## Données et méthode

Le classeur conserve 87 094 commandes historiques, quatre produits, trois fournisseurs et 500 unités de stock de gestion par produit. Les clients sont fictifs. Les opérations de démonstration ne modifient pas rétroactivement le stock à partir des commandes historiques.

La calibration suit le produit choisi parmi les quatre références : elle agrège ses quantités par jour observé et exclut les commandes annulées; les jours absents ne sont pas ajoutés comme des zéros. À titre de repère, P001 comporte 24 515 lignes et 1 006 jours observés. Les fournisseurs alimentent les délais et coûts de commande calibrés. Les valeurs restent ajustables dans Paramètres.

Le moteur utilise une demande normale, une matrice de demandes commune à graine fixe, des calculs VBA en mémoire et au maximum une commande par jour selon la position stock physique + transit. Les comparaisons utilisent R répétitions, des intervalles de Student à 95 % et des différences appariées. La sensibilité mesure l'effet des commandes multiples sur les mêmes demandes.

Les définitions des KPI complémentaires sont dans **ADDENDUM_KPI.md**. Les sources VBA corrigées/exportées sont dans **VBA_Source** ; conserver les fichiers `.frm` et `.frx` ensemble. Les modules concernés sont synchronisés dans le classeur `.xlsm`, y compris les suggestions par produit et le guide. La méthode de calcul des suggestions est détaillée dans **AUDIT_SUGGESTIONS_PRODUITS.md**.

## État de vérification

Consulter **AUDIT_FINAL.md** et **verification.json** pour les tests du 21 octobre 2026. Les tests Excel, le recalcul indépendant et l'actualisation native du modèle Power BI ont été réalisés. Le rendu final et les clics de navigation Power BI restent à confirmer à l'écran : le service de capture a échoué pendant cette intervention.

Les CSV sources originaux et les maquettes cités dans le cahier des charges ne figurent pas dans le dossier reçu. Les données historiques intégrées au classeur ont été conservées et vérifiées ; leur présence ne remplace pas les fichiers originaux pour un dossier de soutenance.
