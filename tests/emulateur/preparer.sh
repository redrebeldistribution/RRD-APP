#!/usr/bin/env bash
# Prépare une copie d'ESSAI de l'application : elle parle au faux bureau, jamais au vrai serveur,
# ni au vrai chat (Firebase), ni aux mises à jour automatiques. À lancer seulement dans une copie jetable
# (GitHub Actions ou un dossier de travail), jamais avant une vraie publication.
#   tests/emulateur/preparer.sh <dossier www> <adresse du faux bureau, ex. http://10.0.2.2:8765>
set -euo pipefail
WWW="$1"; BUREAU="$2"
cat > "$WWW/config.js" <<EOF
/* COPIE D'ESSAI (émulateur) : faux bureau seulement. */
window.RRD_CONFIG = { API_URL: '$BUREAU/exec', FIREBASE_KEY: 'ESSAI-SANS-CHAT', TRANSFERTS_URL: '$BUREAU/exec' };
EOF
sed -i "s#^var MAJ_URL = .*#var MAJ_URL = '$BUREAU/web.json';#" "$WWW/index.html"
sed -i "s#^var FB_URL = .*#var FB_URL = '$BUREAU/firebase/';#" "$WWW/index.html"
# Garde-fou : plus aucune adresse réelle dans la copie d'essai.
if grep -rqE "script\.google\.com/macros|firebaseio\.com|releases/latest/download/web\.json" "$WWW"/*.html "$WWW"/*.js; then
  echo "ERREUR : une adresse réelle est encore dans la copie d'essai" >&2
  grep -rnE "script\.google\.com/macros|firebaseio\.com|releases/latest/download/web\.json" "$WWW"/*.html "$WWW"/*.js >&2
  exit 1
fi
echo "Copie d'essai prête : $WWW -> $BUREAU"
