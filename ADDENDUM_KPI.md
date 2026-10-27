# Évolution des KPI du cahier des charges

Cet addendum documente les onze indicateurs intégrés au simulateur. Le taux de service en volume, le coût total et la règle de sélection des politiques sont conservés. Les unités monétaires sont des USD.

## Simulation : une ligne par scénario et par répétition

Les résultats sont dans `Export_KPI` ; `Synthese!B21:J24` présente les moyennes par scénario. Les paramètres de cette simulation restent dans `Feuille_Controle`.

| N° | Indicateur | Définition | Champ / calcul | Power BI |
|---|---|---|---|---|
| 1 | Disponibilité journalière | Nombre de jours sans unités non servies / N | `TauxDispo` | Pages 1 et 3 |
| 2 | Jours de rupture | Nombre de jours avec au moins une unité non servie | `JoursRupture` | Page 3 |
| 3 | Plus longue rupture | Nombre maximal de jours consécutifs avec des unités non servies | `RuptureMax` | Page 3 |
| 4 | Unités manquantes | Somme des unités demandées et non servies | `UnitesManquantes` | Pages 1 et 3 |
| 5 | Décomposition du coût | Somme des stocks de fin de journée × Cpos ; nombre de commandes × Ccmd ; unités manquantes × Crup | `CoutPossession`, `CoutCommande`, `CoutRupture` | Page 3, graphique empilé |
| 6 | Couverture moyenne théorique | Stock moyen de fin de journée / µ paramétré | `Couverture`, en jours | Page 1 |
| 7 | Part des répétitions respectant l'objectif | Nombre de runs avec service en volume ≥ seuil / nombre de runs du scénario | Mesure DAX `Fiabilite du scenario` | Page 3 |

Le taux de disponibilité journalière représente la proportion de jours simulés sans demande non satisfaite. Il se distingue du taux de service en volume, proportion des unités demandées effectivement servies. Il ne mesure pas les cycles de réapprovisionnement.

Les taux sont stockés entre 0 et 1, puis affichés en pourcentage. Les jours à demande nulle sont des jours sans rupture. Une séquence de rupture en fin de simulation est comptée jusqu'au dernier jour observé. Si µ = 0, la couverture est vide : elle n'est pas remplacée par zéro. Si toute la demande est nulle, le service en volume et la disponibilité valent 100 %.

Le coût total doit être égal à la somme des trois composantes. Les nouveaux KPI ne modifient pas la chronologie de simulation ni les politiques (s,Q). Une rupture n'entraîne pas un report de la demande : la séquence décrit des jours consécutifs de pertes de vente.

La fiabilité utilise le seuil paramétrable de `Feuille_Controle!C18`, exporté dans `Export_Decision[SeuilService]`, initialement 95 %. Elle est calculée sur les répétitions du scénario filtré, et non sur la disponibilité journalière. Une moyenne de service de 95 % ne signifie pas que 95 % des runs respectent ce seuil. Cet indicateur est descriptif et ne constitue pas une garantie de probabilité future.

Sur la page 3, jours de rupture, plus longue rupture et unités manquantes sont les **moyennes des valeurs par run**. La « plus longue rupture moyenne » n'est donc pas le maximum observé sur toutes les répétitions.

`UnitesManquantes` remplace `UnitesRupture` dans `Export_KPI`, sans colonne dupliquée. Le journal quotidien conserve `Rupture` et la table de sensibilité conserve son champ existant `UnitesRupture`.

## Analyse transactionnelle

| N° | Indicateur | Définition retenue | Source / emplacement |
|---|---|---|---|
| 8 | Écart moyen au délai prévu | Moyenne de délai réel − délai prévu ; les valeurs négatives représentent les avances | `T_Commandes`, page 4, par fournisseur |
| 8 complément | Retard positif moyen | Moyenne de MAX(0, délai réel − délai prévu), sur toutes les lignes historiques avec deux délais renseignés | `T_Commandes`, page 4, par fournisseur |
| 9 | Marge moyenne par ligne et part des lignes négatives | Moyenne de `Marge` ; nombre de lignes avec marge < 0 / nombre de lignes avec marge renseignée | `T_Commandes`, page 4 |
| 10 | Produits sous seuil | Nombre de produits avec stock courant strictement inférieur au seuil | `T_Stock`, `Menu!E35` et page 4 |
| 11 | Approvisionnements en attente | Nombre d'approvisionnements au statut En attente ou En transit | `T_Approvisionnements`, `Menu!E36` et page 4 |

Les KPI historiques 8 et 9 utilisent uniquement `Origine = DataCo`, avec les filtres produit et fournisseur actifs. Leur valeur n'est pas fixée à un pourcentage annoncé dans le CDC. La granularité est celle des lignes fournies dans les extraits : aucune moyenne « par commande entière » n'est prétendue sans agrégation vérifiée sur une clé de commande.

Les KPI 10 et 11 suivent la gestion courante. Le filtre produit s'applique au stock et aux approvisionnements ; le filtre fournisseur s'applique aux approvisionnements mais pas au stock, qui n'est pas affecté à un fournisseur unique. Les statuts vides ou reçus ne sont pas comptés. Le classeur crée actuellement le statut `En transit` ; le modèle accepte aussi `EnTransit` pour compatibilité.

Les compteurs du menu sont actualisés après les mouvements de stock, créations et réceptions d'approvisionnement, et à l'export. Power BI affiche le dernier état exporté et actualisé : après une opération VBA, cliquer **Analyse décisionnelle**, puis **Actualiser** dans Power BI.

## Validation

Les contrôles couvrent les valeurs des 90 runs, les trois composantes du coût, les séquences interrompues et finales, l'absence de rupture, µ nul, le stock égal au seuil et le passage d'un approvisionnement en transit à reçu. Le modèle et les définitions du rapport sont contrôlés structurellement. L'exécution des mesures DAX et le rendu restent à vérifier dans Power BI Desktop, absent de l'environnement de réalisation.
