# Comparateur de fichiers Excel

Application locale en français pour comparer des documents administratifs et analyser la qualité de **toutes leurs colonnes**, sans base de données.

## Installation

1. Installez Python 3.10 ou plus récent (cochez « Add Python to PATH » sous Windows).
2. Ouvrez un terminal dans ce dossier et exécutez :

   ```sh
   pip install -r requirements.txt
   ```

3. Lancez l’application :

   ```sh
   streamlit run app.py
   ```

Ouvrez l’adresse locale affichée, normalement http://127.0.0.1:8501. Si les commandes ne sont pas reconnues, utilisez `python -m pip` et `python -m streamlit`. Pour arrêter : Ctrl+C dans le terminal.

## Utilisation

Chargez les fichiers A et B, cliquez sur **Analyser et comparer**, consultez les anomalies, puis cliquez sur **Télécharger le rapport Excel**. Aucun champ n'est à cocher : les identifiants disponibles sont utilisés automatiquement pour retrouver les personnes. Toutes les colonnes des deux fichiers sont contrôlées indépendamment, même pour une personne sans correspondance, et toutes les correspondances fiables sont comparées par défaut.

Trois indicateurs résument l'analyse : **Nombre de lignes analysées**, **Nombre d’erreurs détectées** et **Nombre de lignes sans erreur détectée**. Les lignes sont comptées dans toutes les feuilles analysées des deux fichiers, répétitions comprises. Une ligne ayant plusieurs erreurs n'est retirée qu'une seule fois du nombre de lignes sans erreur. Le nombre d'erreurs correspond aux signalements par cellule (une différence peut signaler une cellule de chaque fichier). Seule la gravité **Erreur** est retenue dans les résultats, les compteurs, les détails et l'export. Les informations et avertissements sont exclus, sans être reclassés en erreurs.

Les deux présentations sont acceptées dans chaque zone. Dans les exemples fournis, `FILE_A.xlsx` contient le nom/prénom et le RIB séparés, et `File_B.xlsx` contient les champs complets : aucune inversion manuelle n’est nécessaire.

Les en-têtes administratifs sont recherchés dans toute la feuille. Les cellules fusionnées et les en-têtes sur plusieurs lignes sont pris en compte. Les feuilles aux intitulés inconnus reçoivent une proposition d'en-têtes explicitement marquée **à vérifier**. **Options avancées** permet de corriger le numéro de ligne pour chaque feuille ou d'exclure une feuille de référence. Les feuilles vides sont ignorées. Plusieurs tableaux administratifs peuvent se succéder dans une feuille, avec des colonnes réordonnées : les en-têtes reconnus déterminent les règles de chaque région. L'analyse reprend après les reports, titres et sous-totaux.

Toutes les colonnes utiles sont conservées, y compris celles avec un titre répété ou absent. Les lignes vides, en-têtes, titres, reports et totaux restent consultables aux coordonnées originales ; seules les lignes `DATA` sont rapprochées comme des personnes et comptées dans les indicateurs de lignes. Les deux feuilles choisies sont rapprochées entre elles ; les autres feuilles bénéficient des contrôles internes et figurent dans le rapport. Pour rapprocher d'autres feuilles, changez la sélection et relancez.

## Analyse de structure

Le traitement suit l'ordre : lecture Excel → classification des régions, lignes et cellules → validation adaptée → comparaison → rapport. `row_classifier.py` combine contenu, identifiants, fusions, styles, position et voisinage. Chaque ligne porte un type, un score heuristique de confiance et une raison. Les régions distinguent données, titres, en-têtes, synthèses, reports, notes, vides et structures inconnues. Les scores ne sont pas des probabilités garanties.

Les règles de dates, montants et identifiants s'appliquent aux données individuelles. Un report tel que `Report... 12.274,00 dhs` est traité comme une synthèse, même dans une colonne habituellement réservée aux dates. Les montants de synthèse acceptent les séparateurs français et les suffixes DH/DHS/MAD ; leurs règles sont séparées des bornes des bénéficiaires. Les formules sont conservées sans recalcul : aucune vérification arithmétique des totaux n'est affirmée en l'absence de règle de calcul. Les valeurs des cellules fusionnées ne sont jamais recopiées automatiquement dans les lignes de données.

La section **Structures à vérifier** présente les lignes incertaines sans les transformer en erreurs de données. Dans **Options avancées** du fichier, ajoutez une correction avec feuille, début, fin et type (`DATA`, `REPORT`, `TITLE`, etc.), puis cliquez à nouveau sur **Analyser et comparer**. La correction invalide le rapport précédent et déclenche les validations du type choisi. Les plages ne doivent pas se chevaucher et une ligne contenant des valeurs ne peut pas être déclarée vide. Une structure inhabituelle peut nécessiter cette vérification humaine, notamment des tableaux aux en-têtes inconnus ou disposés côte à côte.

Le visualiseur indique le type de ligne sans colorer les reports comme des erreurs. Le détail d'une erreur précise la règle et son contexte d'application. L'export ajoute **Classification des lignes**, **Régions détectées** et **Rôles des cellules**, avec les coordonnées et raisons ; les feuilles sources annotées conservent aussi les lignes de présentation.

## Visualiseur et tableau de bord

La section **Personnes absentes d’un fichier** présente deux listes : personnes présentes dans A et absentes de B, puis personnes présentes dans B et absentes de A. Chaque ligne indique la feuille, le numéro de ligne Excel, le nom et les identifiants disponibles. Le rapprochement utilise les feuilles sélectionnées et les mêmes règles d'identification que la comparaison ; une identité incomplète ou une correspondance possible/ambiguë n'est pas déclarée absente. Les reports et titres sont exclus, les lignes répétées sont conservées. Les deux listes figurent aussi dans le rapport Excel, dans **Dans A absentes de B** et **Dans B absentes de A**.

La liste des erreurs indique le fichier, la feuille, la ligne Excel, la personne, la colonne, le type et la description. Cliquez sur une ligne pour afficher la cellule concernée en première colonne de l'extrait, avec son surlignage et les valeurs comparées disponibles. La recherche **Rechercher une personne ou une valeur** et les filtres **Fichier**, **Colonne** et **Type d'anomalie** permettent de retrouver un problème. Les commandes **Précédente / Suivante** restent disponibles. Aucun filtre de gravité n'est nécessaire.

Le visualiseur conserve toutes les colonnes d'origine avec la lettre Excel et le numéro de ligne, par pages de 100 lignes. **Seulement les lignes avec anomalies** réduit l'affichage si nécessaire. Une vue côte à côte permet de consulter les lignes rapprochées. Les anciennes catégories de résultats, les compteurs détaillés et les colonnes non comparées restent accessibles dans **Détails de l’analyse**.

- **Erreur** (rouge) : règle configurée violée ou différence confirmée. Une différence ne dit pas quel fichier est correct.
- Les valeurs simplement suspectes, les formules non calculées et les variations de présentation ne sont pas signalées comme erreurs. Les cellules correspondantes ne reçoivent ni couleur ni commentaire d'anomalie dans l'export.
- Les protections contre les rapprochements incertains restent actives : l'absence d'erreur détectée ne signifie pas qu'une correspondance a été confirmée.

Une anomalie correspond à une règle signalée dans une cellule. Plusieurs règles peuvent concerner la même cellule. Le compteur **Cellules signalées (distinctes)** évite ce cumul. Une différence correspond à une paire de valeurs d'un champ sur une ligne rapprochée : les deux cellules mises en couleur ne sont pas comptées comme deux différences. La détection des doublons est désactivée : les valeurs répétées ne constituent pas une anomalie et aucun contrôle d'unicité n'est proposé.

## Configuration avancée

La configuration est facultative et destinée aux personnes connaissant les règles métier ; il ne s'agit pas d'un espace protégé par authentification. Elle reste en mémoire dans la session et ses valeurs sont incluses dans l'export. Toute modification des données ou de la configuration invalide le résultat précédent et nécessite une nouvelle comparaison.

**Correspondances :** les alias connus (DDP/DDR, CIN, début/fin), les noms et RIB reconstitués ainsi que les intitulés exactement identiques sont proposés. Aucune correspondance n'est créée sur une simple ressemblance de noms. Les intitulés ambigus ne sont pas rapprochés automatiquement. **NET(DH) / Montant net** reste inactif tant que l'administrateur n'a pas confirmé qu'il s'agit de la même valeur métier. Le tableau permet d'activer, désactiver, ajouter, supprimer et corriger les correspondances. Les colonnes sans correspondance restent analysées en interne.

**Règles par colonne :** obligatoire/facultatif, type (`auto`, `text`, `number`, `date`, `identifier`), longueurs minimale/maximale, valeurs autorisées séparées par `|`, bornes numériques et dates. Une borne vide n'impose aucune limite. Les règles personnalisées acceptent `prefixe:texte`, `suffixe:texte`, `contient:texte` ou `sans:texte` ; elles ne permettent jamais d'exécuter du Python.

Les identifiants et noms connus sont obligatoires par défaut ; ce choix est modifiable. Les CIN et DDP/DDR répétés sont acceptés. Les dates et colonnes inconnues sont facultatives. Les montants négatifs sont autorisés sauf minimum explicitement configuré. Aucun catalogue de catégories métier n'est inventé. Pour une colonne inconnue, une tendance forte des valeurs peut suggérer un type ; les écarts restent des avertissements tant que le type n'est pas imposé.

Les contrôles comprennent les valeurs manquantes obligatoires, les types, les dates impossibles et dates de fin antérieures au début, les espaces, les variantes de casse/écriture, les caractères inhabituels, les formats RIB et les règles configurées. Les valeurs numériques très éloignées des autres sont signalées comme avertissements lorsque la colonne contient suffisamment de nombres et une dispersion mesurable.

## Règles de comparaison

- Les champs nom, CIN, DDP/DDR et RIB disponibles servent automatiquement à retrouver les personnes, sans cases à cocher. Les correspondances actives déterminent ensuite les comparaisons cellule par cellule, y compris pour les colonnes supplémentaires.
- CIN, DDP/DDR et RIB fiables sont recherchés dans B. Si plusieurs lignes B peuvent correspondre, le résultat est **Correspondance ambiguë** pour éviter de choisir arbitrairement une personne. Si aucun identifiant ne correspond, le nom sélectionné sert de recours. Les noms sont comparés sans accents, sans différence de casse ni espaces superflus ; les originaux sont conservés pour l’affichage.
- Plusieurs lignes A peuvent être comparées à la même ligne B. Une répétition dans A ne crée aucune anomalie ; chaque ligne est traitée indépendamment.
- Sinon, une information sélectionnée manquante ou incertaine donne **Incomplet**, même si une différence certaine existe aussi. Les contrôles détaillés affichent ces différences.
- Sans candidat, une ligne complète donne **Absent**. Avec un candidat unique, elle donne **Différence** ou **Correspondance** selon les correspondances actives. Sans identifiant commun sélectionné, seule l'analyse interne est effectuée et les lignes sont marquées **Incomplet**.
- Si tous les identifiants disponibles ont changé, le logiciel ne peut pas deviner la personne : elle sera absente. Si seul le RIB est disponible, un changement de RIB peut empêcher le rapprochement. Il n’y a pas de rapprochement approximatif des noms.

## Identifiants et RIB

Les identifiants sont conservés comme texte. Les formats Excel simples `000000` sont respectés. Le suffixe `.0` est retiré pour les identifiants entiers, sans conversion flottante de chaînes. Le RIB marocain est constitué de **3 chiffres banque + 3 chiffres ville + 16 chiffres compte + 2 chiffres clé**, soit **24 chiffres**. Les composants sont complétés à gauche avec des zéros ; un composant manquant ne devient jamais zéro. Le RIB complet n’est jamais converti en nombre. Les RIB mal formés sont à vérifier (aucun contrôle bancaire de validité de la clé n’est effectué).

Un identifiant stocké comme nombre Excel de plus de 15 chiffres est marqué comme incertain et exclu du rapprochement. Des chiffres déjà perdus par Excel ne sont pas récupérables : consultez la source originale. Les formules dans les champs comparés sont à vérifier ; enregistrez leurs résultats comme valeurs. Les formats numériques personnalisés complexes ne sont pas reconstitués : stockez les identifiants comme texte. Les RIB séparés affichés dans les résultats sont les RIB reconstitués.

Les dates textuelles sont interprétées selon les formats français jour/mois/année et ISO année-mois-jour. Les nombres acceptent une virgule ou un point décimal et des espaces de groupement ; les écritures ambiguës sont à vérifier. Les colonnes financières ne sont pas arrondies pour la comparaison. Les valeurs d'origine sont toujours affichées avec leurs coordonnées.

## Rapport Excel annoté

Le téléchargement contient **Résumé global**, **Anomalies fichier A**, **Anomalies fichier B**, **Différences entre fichiers**, **Correspondances ambiguës**, **Colonnes non comparées**, **Fichier A annoté** et **Fichier B annoté**. Les autres catégories de résultats sont conservées ; les autres feuilles analysées reçoivent des copies annotées supplémentaires. **Périmètre analysé** relie chaque copie à sa feuille d'origine. Les règles et correspondances sont également incluses. Aucun onglet ni compteur de doublons n'est généré.

Les copies annotées conservent les coordonnées, valeurs, fusions, polices, bordures, alignements, formats numériques et dimensions des cellules sources. Une couleur et un commentaire expliquent chaque cellule signalée. Les formules sont conservées sous forme de texte non exécutable, sans recalcul. Les graphiques, images, mises en forme conditionnelles et objets Excel avancés ne sont pas reproduits ; ce rapport est une copie d'analyse, pas une sauvegarde intégrale du classeur. Les fichiers sources ne sont jamais écrasés.

## Confidentialité

Les classeurs et exports restent en mémoire dans la session locale. Aucun envoi à une API, aucune télémétrie et aucun enregistrement automatique des fichiers téléversés. Le téléchargement est enregistré uniquement à votre demande. Le serveur écoute sur `127.0.0.1` et accepte des fichiers jusqu’à 25 Mo ; réservez-le à des classeurs locaux de confiance. Fermez la session et arrêtez le serveur après utilisation. Les exports conservent les identifiants comme texte et neutralisent les formules provenant des valeurs sources.

## Vérification et organisation

```sh
python -m pytest -q
```

`app.py` gère le parcours ; `excel_reader.py` conserve les tableaux et coordonnées ; `normalizer.py` et `comparator.py` rapprochent les personnes ; `data_quality.py` applique les règles ; `field_mapping.py` compare les champs équivalents ; `quality_report.py` produit les compteurs ; `quality_ui.py` fournit la configuration et le visualiseur ; `exporter.py` génère le rapport annoté.

Les tests couvrent les règles, les correspondances simples et composées, les colonnes inconnues, les feuilles supplémentaires, les exports, les interactions Streamlit et un cas de 1 500 lignes par fichier (33 000 cellules). Les exemples locaux sont testés en mémoire sans enregistrer de copie ni journaliser leurs données personnelles. La mise en mémoire et la pagination conviennent à des centaines ou milliers de lignes ; les classeurs très volumineux avec beaucoup de formats ou de cellules fusionnées peuvent demander davantage de ressources.
