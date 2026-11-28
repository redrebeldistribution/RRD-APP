#!/usr/bin/env bash
# Lancé dans l'émulateur de GitHub Actions : installe l'APK d'essai, joue les scénarios, garde le journal Android.
set -uo pipefail
cd "$(dirname "$0")"
OUT="${1:-captures}"; mkdir -p "$OUT"
adb wait-for-device
adb shell settings put global window_animation_scale 0
adb shell settings put global transition_animation_scale 0
adb shell settings put global animator_duration_scale 0
adb install -r -g ../../android/app/build/outputs/apk/debug/app-debug.apk || exit 1
adb logcat -c
adb logcat -v time > "$OUT/journal-android.txt" 2>&1 &
LOGCAT=$!
python3 scenarios.py emulateur http://127.0.0.1:8765 "$OUT"; RES=$?
sleep 2; kill $LOGCAT 2>/dev/null
# Plantage de l'application pendant les essais = échec, même si les scénarios sont passés.
if grep -q "FATAL EXCEPTION" "$OUT/journal-android.txt" && grep -A3 "FATAL EXCEPTION" "$OUT/journal-android.txt" | grep -q "com.redrebel.rrdapp"; then
  echo "❌ L'application a planté (FATAL EXCEPTION) : voir journal-android.txt"
  printf '\n**❌ Plantage de l application (FATAL EXCEPTION)** : voir journal-android.txt\n' >> "$OUT/resume.md"
  RES=1
fi
exit $RES
