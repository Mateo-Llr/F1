# F1 — Saison

Site statique pour consulter les saisons, pilotes, écuries et courses de Formule 1.

## Publication sur GitHub Pages

Le site charge automatiquement `datas.xlsx` depuis la racine du dépôt. Publiez le contenu du dépôt avec GitHub Pages depuis la branche `main`, dossier `/ (root)`. Aucun serveur ou compilation n’est nécessaire.

Pour mettre à jour les données, remplacez `datas.xlsx` en conservant les noms d’onglets attendus par l’application, puis publiez le changement.

## Frise chronologique des écuries

L’onglet `Historique_Écuries` contient les saisons des constructeurs engagés de 1950 à 1999, y compris les Indianapolis 500 comptant pour le championnat. Les écuries sont regroupées par constructeur : plusieurs entrants utilisant le même constructeur ne créent donc pas de lignes distinctes dans la frise. L’onglet `Conducteurs` et les saisons d’écuries incluent aussi les pilotes inscrits dans ces engagements. Ces données historiques proviennent des listes saisonnières et des fiches pilotes de [F1DB](https://github.com/f1db/f1db), sous licence CC BY 4.0.

## Images des pilotes et des écuries

Les années des images disponibles sont indexées dans `script.js` (`DRIVER_IMAGE_YEARS` et `CAR_IMAGE_YEARS`). Lorsqu’une image n’existe pas pour la saison demandée, le site choisit directement l’année disponible la plus proche au lieu de tester les fichiers année par année. Mettez à jour l’index correspondant lorsque vous ajoutez ou retirez des images dans `assets/drivers` ou `assets/cars`.
