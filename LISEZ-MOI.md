# BCF Classement — Version démo

Complément Outlook de classement assisté des e-mails, pour le Bureau comptable et fiscal Jacques Rozenblum.
Cette version est une **démonstration** : les dossiers, suggestions et le message affiché sont **fictifs**, et **aucune action ne touche votre boîte mail réelle**. Son but est de vérifier que le complément s'installe, s'héberge et s'affiche correctement dans votre New Outlook, avant le branchement sur vos vrais dossiers (étape suivante, après l'inscription Entra ID).

## Contenu du dossier

- `manifest.xml` — le fichier à charger dans Outlook.
- `taskpane.html` — l'interface du volet (avec votre logo).
- `taskpane.js` — la logique (suggestions, recherche, favoris, classement, annulation).
- `assets/` — votre logo et les icônes du complément.

## Étape 1 — Héberger les fichiers (gratuit)

Le complément est une petite application web : ses fichiers doivent être accessibles via une adresse `https`. Solution gratuite recommandée : **GitHub Pages**.

1. Créez un compte sur https://github.com (gratuit) si vous n'en avez pas.
2. Créez un nouveau dépôt (Repository), par exemple `bcf-classement`, en **Public**.
3. Téléversez-y tout le contenu de ce dossier (glisser-déposer les fichiers + le dossier `assets`).
4. Dans le dépôt : onglet **Settings** → **Pages** → sous « Branch », choisissez `main` puis `/ (root)` → **Save**.
5. Au bout d'une minute, GitHub affiche l'adresse publique, du type :
   `https://VOTRE-NOM.github.io/bcf-classement/`
   C'est votre **URL de base**.

## Étape 2 — Renseigner l'URL dans le manifeste

Dans `manifest.xml`, remplacez **toutes** les occurrences de `URL_DE_BASE` par votre URL de base **sans la barre oblique finale**.

Exemple : si votre URL est `https://jrozenblum.github.io/bcf-classement/`,
alors `URL_DE_BASE/taskpane.html` devient
`https://jrozenblum.github.io/bcf-classement/taskpane.html`.

Renseignez aussi un identifiant unique : remplacez la ligne
`<Id>11111111-2222-3333-4444-555555555555</Id>`
par un nouvel identifiant (générez-en un sur https://www.guidgenerator.com).

Re-téléversez le `manifest.xml` modifié sur GitHub (ou modifiez-le directement en ligne).

## Étape 3 — Charger le complément dans Outlook

1. Dans votre navigateur, allez sur **https://aka.ms/olksideload** (la page testée précédemment).
2. Connectez-vous, attendez l'ouverture de la fenêtre « Compléments ».
3. Allez dans **Mes compléments** → **Ajouter un complément personnalisé** → **Ajouter à partir d'un fichier**.
4. Sélectionnez votre `manifest.xml` (la version avec votre URL renseignée).
5. Confirmez. Le complément « BCF Classement » apparaît.

## Étape 4 — Utiliser la démo

1. Ouvrez New Outlook, sélectionnez n'importe quel message reçu.
2. Dans le ruban, cliquez sur le bouton **Classer** (groupe « BCF Classement »).
3. Le volet s'ouvre à droite, avec votre logo et l'interface de démonstration.
4. Essayez : cliquer un dossier suggéré, taper dans la recherche puis Entrée, cliquer un favori, tester « Annuler ».

Rappel : tout est simulé à ce stade. Rien n'est déplacé dans votre boîte.

## En cas de souci

- **Le volet reste blanc** : vérifiez que l'URL de base est correcte dans le manifeste et que les fichiers sont bien en ligne (ouvrez `URL_DE_BASE/taskpane.html` dans un navigateur — la démo doit s'afficher).
- **Le bouton n'apparaît pas** : fermez puis rouvrez Outlook ; le chargement d'un complément peut prendre une minute.
- **Le logo ne s'affiche pas** : vérifiez que le dossier `assets` a bien été téléversé avec ses images.

## Prochaine étape

Une fois la démo confirmée chez vous, on remplacera la couche de données fictive par les appels à **Microsoft Graph** (vos vrais dossiers et le déplacement réel des messages), en s'appuyant sur l'inscription Entra ID. Seul le fichier `taskpane.js` (partie « COUCHE DONNÉES ») sera modifié ; l'interface restera identique.
