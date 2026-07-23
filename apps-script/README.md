# Backend LustuZone — projet Apps Script découpé en modules

Ce dossier remplace l'ancien fichier unique `apps-script.js` par 17 fichiers,
un par module fonctionnel. Le fichier `apps-script.js` original n'a pas été
touché et reste utilisable tel quel tant que vous n'avez pas basculé dessus.

## Installer cette version (première fois)

1. Ouvre ton Google Sheet existant > Extensions > Apps Script.
2. **Ne supprime rien pour l'instant** si tu veux garder un filet de sécurité :
   soit tu testes d'abord sur une copie du classeur, soit tu notes que tu
   pourras revenir à l'ancien `apps-script.js` (gardé dans ce dossier du
   projet) en cas de souci.
3. Supprime le contenu du fichier `Code.gs` par défaut de l'éditeur (ou
   renomme-le/réutilise-le pour un des fichiers ci-dessous).
4. Pour chacun des 17 fichiers de ce dossier, dans l'éditeur Apps Script :
   fichier (icône +) > Script, donne-lui **exactement** le même nom (sans
   l'extension `.gs`, Apps Script l'ajoute automatiquement), et colle le
   contenu correspondant.
   - Config.gs, Router.gs, Sync.gs, Setup.gs, Auth.gs, Grid.gs, Presences.gs,
     Paiements.gs, Evenements.gs, Covoiturage.gs, Osteo.gs, Actualites.gs,
     Support.gs, Salaries.gs, Photos.gs, Notifications.gs, Backup.gs
5. Sélectionne la fonction `setup` en haut, clique "Exécuter" (▶) — comme
   avant, ça initialise les feuilles manquantes sans toucher aux données déjà
   présentes.
6. Si tu avais déjà une Application Web déployée, pas besoin de changer
   l'URL : "Déployer > Gérer les déploiements > crayon > Nouvelle version >
   Déployer" suffit pour appliquer ce code.

## Mettre à jour après ce découpage

Tu n'as plus besoin de tout recoller : identifie le fichier concerné par ton
changement (ex. une modif sur les RDV ostéo → `Osteo.gs` uniquement) et colle
seulement celui-là dans l'éditeur, puis redéploie une nouvelle version.

## Réutiliser pour un autre club

Voir la section "Workflow de réutilisation" du plan de restructuration :
copier ce dossier, éditer `Config.gs` (barème, effectif, nom du club, email
de support), supprimer les fichiers des modules non désirés (ex. `Osteo.gs`
+ `Salaries.gs`), et retirer les lignes correspondantes dans
`Router.gs` (objet `API_HANDLERS`) et `Setup.gs` (fonction `setup()`).
