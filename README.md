# RRD – Application Android (Rapport de fermeture)

Application Android hors connexion pour les stations RRD (rapports de fermeture, photos, transferts, camions, ventes).

- `www/` : l'application (écrans de l'application, avec enregistrement hors connexion).
- `www/config.js` : adresse du serveur (`.../exec`).
- `apps-script/Mobile.gs` : fichier à ajouter au projet Apps Script « RRD – Rapport de fermeture ».
- `android/` : projet Android (Capacitor). Le fichier .apk est construit automatiquement par GitHub Actions à chaque modification, et publié dans « Releases ».

Le mot de passe du certificat de signature est dans le secret `RRD_KEYSTORE_PASSWORD` du dépôt. Ne pas supprimer `android/keystore/rrd-release.jks` : sans lui, les téléphones ne pourraient plus installer les mises à jour.
