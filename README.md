# Comparateur de fichiers Excel

Application locale en français pour comparer des documents administratifs, sans base de données.

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

Chargez les fichiers A et B, sélectionnez les champs, cliquez sur **Comparer les fichiers**, puis téléchargez le résultat Excel. Tous les champs sont cochés au départ. Les résultats sont calculés pour chaque ligne de A ; les personnes uniquement présentes dans B ne sont pas listées.

Les deux présentations sont acceptées dans chaque zone. Dans les exemples fournis, `FILE_A.xlsx` contient le nom/prénom et le RIB séparés, et `File_B.xlsx` contient les champs complets : aucune inversion manuelle n’est nécessaire.

Les 50 premières lignes de chaque feuille sont examinées. Les cellules fusionnées sont prises en compte, y compris les en-têtes répartis sur plusieurs lignes. S’il existe plusieurs feuilles compatibles, choisissez la bonne feuille. En cas d’échec, **Options avancées** permet de choisir la feuille et la ligne d’en-têtes. Les titres de colonnes doivent correspondre aux variantes reconnues dans `excel_reader.py` ; aucun mappage technique n’est demandé aux utilisateurs. Les lignes sans valeur dans les colonnes utiles et les en-têtes répétés sont ignorés ; les colonnes étrangères à la comparaison ne sont pas chargées dans les résultats.

## Règles de comparaison

- Seuls les champs sélectionnés servent au rapprochement et aux contrôles.
- CIN, DDP/DDR et RIB fiables sont recherchés dans B. Si plusieurs identifiants désignent des lignes différentes, le résultat est **Doublon**. Si aucun identifiant ne correspond, le nom sélectionné sert de recours. Les noms sont comparés sans accents, sans différence de casse ni espaces superflus ; les originaux sont conservés pour l’affichage.
- Une même ligne B revendiquée par plusieurs lignes A, plusieurs candidats B, ou des lignes A identiques sur tous les champs choisis donnent **Doublon**. Aucun candidat n’est choisi arbitrairement. Les numéros de lignes permettent de vérifier les sources.
- Sinon, une information sélectionnée manquante ou incertaine donne **Incomplet**, même si une différence certaine existe aussi. Les contrôles détaillés affichent ces différences.
- Sans candidat, une ligne complète donne **Absent**. Avec un candidat unique, elle donne **Différence** ou **Correspondance** selon les champs sélectionnés.
- Si tous les identifiants sélectionnés ont changé, le logiciel ne peut pas deviner la personne : elle sera absente. Sélectionner uniquement le RIB ne permet pas d’identifier une personne dont le RIB a changé. Il n’y a pas de rapprochement approximatif des noms.

## Identifiants et RIB

Les identifiants sont conservés comme texte. Les formats Excel simples `000000` sont respectés. Le suffixe `.0` est retiré pour les identifiants entiers, sans conversion flottante de chaînes. Le RIB marocain est constitué de **3 chiffres banque + 3 chiffres ville + 16 chiffres compte + 2 chiffres clé**, soit **24 chiffres**. Les composants sont complétés à gauche avec des zéros ; un composant manquant ne devient jamais zéro. Le RIB complet n’est jamais converti en nombre. Les RIB mal formés sont à vérifier (aucun contrôle bancaire de validité de la clé n’est effectué).

Un identifiant stocké comme nombre Excel de plus de 15 chiffres est marqué comme incertain et exclu du rapprochement. Des chiffres déjà perdus par Excel ne sont pas récupérables : consultez la source originale. Les formules dans les champs comparés sont à vérifier ; enregistrez leurs résultats comme valeurs. Les formats numériques personnalisés complexes ne sont pas reconstitués : stockez les identifiants comme texte. Les RIB séparés affichés dans les résultats sont les RIB reconstitués.

## Confidentialité

Les classeurs et exports restent en mémoire dans la session locale. Aucun envoi à une API, aucune télémétrie et aucun enregistrement automatique des fichiers téléversés. Le téléchargement est enregistré uniquement à votre demande. Le serveur écoute sur `127.0.0.1` et accepte des fichiers jusqu’à 25 Mo ; réservez-le à des classeurs locaux de confiance. Fermez la session et arrêtez le serveur après utilisation. Les exports conservent les identifiants comme texte et neutralisent les formules provenant des valeurs sources.

## Vérification et organisation

```sh
python -m pytest -q
```

`app.py` contient l’interface ; `excel_reader.py` détecte et lit les tableaux ; `normalizer.py` normalise les valeurs ; `comparator.py` rapproche les lignes ; `exporter.py` crée les six onglets Excel. Les tests utilisent des données fictives ; un test facultatif contrôle seulement la structure et le nombre de lignes des deux exemples locaux.
