# Audit DataCo Global

21 octobre 2026. Référence : **Cahier_des_charges.pdf**, version V4.

## Conclusion

Le volet transactionnel et le moteur de simulation satisfont les besoins testés du cahier des charges. Les erreurs signalées de comparaison et d'actualisation ont été reproduites puis corrigées. Les rapports du formulaire présentent des données et des colonnes cohérentes. L'actualisation native du modèle Power BI a réussi avec les exports locaux.

La validation visuelle complète du formulaire et des pages Power BI courantes, ainsi que les clics sur les boutons Power BI, ne sont pas certifiés : le service de capture Windows renvoyait `window capture timed out` et l'entrée par élément renvoyait `coordinate input geometry is unavailable`. Les tests du formulaire exécutent les événements VBA dans Excel natif ; ils ne constituent pas une inspection de captures. Les destinations Power BI ont été contrôlées structurellement et contre les schémas Microsoft.

## Corrections

| Problème | Correction |
|---|---|
| Erreur 13 « Incompatibilité de type » pendant la comparaison ou Actualiser | La fenêtre du classeur était enregistrée masquée, ce qui faisait échouer la lecture de `Application.Calculation`. Restauration explicite de la visibilité avant les opérations. Régression testée en masquant volontairement la fenêtre. |
| Excel inaccessible à la fermeture du formulaire | Fermeture et croix restaurent le classeur sur Menu. Les feuilles de gestion restent visibles. |
| Rapport vide ou minuscule | Dimensionnement explicite des contrôles et remise en page après l'affichage effectif de la UserForm. Vue synthèse renseignée dès l'accès à Rapports. |
| Colonnes de résultats décalées | Stock moyen, ruptures et commandes sont reliés à leurs indicateurs corrects. Formats indépendants de la largeur des cellules Excel. |
| Rapport peu exploitable | Six vues, en-têtes adaptatifs, contexte de calcul, état périmé et avertissement pour résultats personnalisés ne correspondant plus aux paramètres. |
| Défilement horizontal parasite | Suppression sur le formulaire et ses pages. Largeurs des colonnes ajustées à la liste ; défilement vertical conservé pour les journaux. |
| Rechargement des paramètres | Suppression de la recalibration involontaire lors du rechargement. Le changement de produit recharge aussi les champs L, Ccmd et Crup. |
| État Excel après erreur | Restauration du calcul, des événements et du rafraîchissement d'écran si une comparaison échoue. Les résultats restent marqués périmés. |
| Source Power BI sur un ancien ordinateur | Paramètre TMDL FolderExports réécrit selon le dossier courant. Script Configurer_PowerBI mis à jour pour le modèle TMDL livré. Écriture UTF-8 sans BOM, exigée par Power BI et contrôlée après le bouton Excel. |
| Navigation Power BI (version auditée initialement) | 36 boutons PageNavigation avec une destination existante explicite sur les six pages. Le rapport courant a depuis été réduit à quatre pages et sa navigation compte 16 boutons. |
| Accueil et rangement | Menu en premier avec Ouvrir le formulaire et raccourcis. Classeur, sources VBA, exports et projet Power BI regroupés à la racine. |

## Couverture fonctionnelle

| Exigence | Vérification et résultat |
|---|---|
| BF-T1 Produits | Ajout, modification, consultation et suppression autorisée ; suppression d'une référence utilisée refusée. PASS natif. |
| BF-T2 Fournisseurs | CRUD et prise en compte des paramètres fournisseurs dans la calibration. PASS natif. |
| BF-T3 Commandes | Vente partielle, unités non servies, annulation, remise en stock des seules unités vendues, double annulation refusée. PASS natif. |
| BF-T4 Stocks | Mouvements, seuil, alerte sous seuil et stock négatif refusé. PASS natif. |
| BF-T5 Approvisionnements | Création et réception mettant à jour le stock une seule fois ; double réception refusée. PASS natif. |
| BF-T6 Persistance | Tables structurées du classeur et exports transactionnels inchangés après intervention. Aucun enregistrement de test dans le livrable. |
| BF-T7 Passerelle | Export CSV UTF-8, source locale configurable et lancement Power BI par le bouton du formulaire. Voir résultat précis dans verification.json. |
| BF-D1 Paramètres | Validation des bornes, R >= 2, L >= 1, rechargement sans altération du mu saisi, sélection d'un produit. PASS natif. |
| BF-D2 Demande | Box-Muller, arrondi et borne zéro, graine reproductible, dix tirages et matrice commune. PASS natif et calcul indépendant. |
| BF-D3 Moteur | Réception, vente, position de stock, au plus une commande par jour, réception t + L ; calculs en tableaux mémoire. PASS sur cas déterministes et journaux complets. |
| BF-D4 Ruptures | Quantification en unités, jours et séquences ; cas de demande nulle. PASS natif et indépendant. |
| BF-D5 KPI | Concordance des 90 runs, décomposition des coûts et couverture. PASS indépendant. |
| BF-D6 Comparaison | Trois politiques sur les mêmes demandes, moyennes, écarts-types, Student et différences appariées. PASS indépendant. |
| BF-D7 Exports | 1 095 lignes du run 1 ; 90 lignes KPI ; 180 résultats de sensibilité. PASS. |
| BF-D8 Tableaux de bord (version auditée initialement) | Six pages alors définies ; modèle chargé et actualisé nativement ; requêtes DAX contrôlées. Le rapport courant a depuis été réduit à quatre pages; le rendu final et les clics restent à confirmer visuellement. |
| BF-D9 Sensibilité | Oui/Non sur la même demande pour les trois scénarios. Concordance des 180 runs. PASS. |

## Données et calculs

- 87 094 commandes historiques conservées, quatre produits, trois fournisseurs.
- P001 : 24 515 lignes et 1 006 jours observés. Moyenne journalière 73,2584493042 ; écart-type échantillon 16,7168042835 ; coût de rupture moyen 6,7137500714 USD/unité. Recalculés indépendamment depuis l'historique.
- Stocks de gestion : quatre lignes, 500 unités par produit. Aucun approvisionnement ou mouvement de test conservé.
- Comparaison indépendante de toutes les colonnes des 90 runs KPI, des 1 095 lignes quotidiennes et des 180 runs de sensibilité. Tolérance numérique relative 1e-8 ; identifiants et journaux entiers comparés exactement.
- Vérification des intervalles de coût et des trois différences appariées. Le quantile Student a été recalculé indépendamment par intégration numérique.
- 107 fichiers JSON Power BI contrôlés ; 36 destinations de navigation existantes. Les fichiers ayant un `$schema` ont été validés contre le schéma Microsoft correspondant.
- Actualisation native du modèle Power BI puis lecture DAX : 87 094 commandes, 90 KPI et 1 095 lignes quotidiennes. Les résultats par scénario concordent avec Excel.
- Rendus PDF natifs de Menu, Synthese et Evolution_Stock inspectés : boutons et tableaux lisibles ; courbe du stock avec trois séries et légende.

| Scénario | Coût moyen USD | Service moyen | Stock moyen | Ruptures moyennes en unités |
|---|---:|---:|---:|---:|
| Prudent | 3 891,07 | 99,9968 % | 281,72 | 0,87 |
| Flux tendu | 4 890,20 | 98,8056 % | 107,34 | 319,17 |
| Équilibré | 3 282,05 | 99,9416 % | 193,09 | 15,63 |

À paramètres par défaut et seuil de service de 95 %, Équilibré minimise le coût moyen parmi les politiques admissibles. Cette conclusion est calculée et dépend des paramètres ; elle n'est pas imposée dans le code.

## Réserves pour la soutenance

1. Vérifier à l'écran les quatre pages courantes et les boutons Power BI. Le modèle et ses données ont été testés avant le retrait des pages, mais le service de capture n'a pas permis de certifier ce parcours visuel. Dans Desktop en édition, Ctrl + clic est nécessaire pour exécuter une action de navigation.
2. Les fichiers sources originaux `DataCo_Cleaned_Orders`, `Dataset_4_Produits_DataCo.csv`, `Dataset_1_Produit_RipDeck.csv` et les maquettes cités dans le cahier des charges n'étaient pas présents dans le dossier reçu. Les données intégrées sont conservées et la calibration est vérifiée ; la provenance complète jusqu'au fichier brut ne peut pas être certifiée sans ces sources.
3. La sélection des quatre produits est une extension : chaque simulation reste mono-produit et le réglage livré utilise P001, conformément au cas Rip Deck de référence. Les simulations ne représentent pas une chaîne multi-produits simultanée.
4. Les limites de modélisation du cahier des charges restent applicables : délai fixe, coût de possession supposé, marge moyenne pertes incluses et clients fictifs.

Les détails techniques des contrôles réalisés et le nettoyage sont consignés dans **verification.json**.

## Complément du 21 octobre 2026 : sélection du produit et boutons

Les sources VBA exportées ont été ajustées pour afficher le code et le nom du dernier produit dans Power BI, écrire `Simulation_Contexte.csv` à chaque export, conserver le produit sélectionné lors de **Réinitialiser**, actualiser l'onglet Guide et faire fonctionner **Voir le guide Excel** en activant cet onglet. Les quatre produits déjà présents au catalogue restent simulés un par un; un changement de produit marque les résultats précédents comme périmés.

Le modèle PBIP définit une table de contexte indépendante des filtres de scénario. L'identité du produit simulé est affichée sur chacune des quatre pages. Les 64 fichiers JSON du rapport se chargent, les 16 boutons de navigation pointent vers les quatre pages existantes, et les visuels sont dans les limites de leurs pages. Cette vérification est structurelle : Power BI Desktop n'a pas été relancé pour actualiser ou rendre les changements.

Les corrections VBA sont présentes dans `VBA_Source`, mais n'ont pas été réinjectées dans `DataCo_Logistique.xlsm` pendant cette intervention : l'activation de la fenêtre Excel a échoué. Il faut importer/synchroniser les sources avant que les changements de macros et du guide deviennent actifs dans le classeur. La cartographie et les contrôles statiques sont décrits dans **AUDIT_BOUTONS.md**. Les tests natifs déjà consignés plus haut restent les vérifications antérieures, et non une nouvelle exécution aujourd'hui.

La version courante du rapport ne contient plus que **Vue globale**, **Stock**, **Scénarios** et **Coûts**. Les pages **Fournisseurs** et **Transactions**, ainsi que leurs boutons de navigation, ont été supprimées. La calibration documentée et l'indicateur de produit sur les quatre pages Power BI utilisent le produit associé à la dernière simulation à jour; la mention spécifique à Rip Deck a été retirée du texte d'aide VBA. Les anciennes observations chiffrées sur six pages ci-dessus décrivent la version contrôlée avant ce changement de périmètre.

Références techniques : [navigation Power BI](https://learn.microsoft.com/en-us/power-bi/create-reports/button-navigators), [propriété Excel Calculation](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculation), [actions PageNavigation PBIR](https://github.com/microsoft/skills-for-fabric/blob/main/skills/powerbi-report-cli/references/authoring/button-part-02.md).
