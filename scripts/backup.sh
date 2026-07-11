#!/usr/bin/env bash
#
# backup.sh — Sauvegarde l'intégralité des données DevisFlow
#
# Génère une archive horodatée dans backups/
# À programmer en cron pour une exécution quotidienne.

set -e

cd "$(dirname "$0")/.."
PROJECT_ROOT="$(pwd)"

BACKUP_DIR="${PROJECT_ROOT}/backups"
mkdir -p "$BACKUP_DIR"

TIMESTAMP=$(date +%Y%m%d-%H%M%S)
ARCHIVE="$BACKUP_DIR/devisflow-$TIMESTAMP.tar.gz"

echo "→ Sauvegarde DevisFlow"
echo "  Source     : $PROJECT_ROOT/backend/storage/"
echo "  Destination: $ARCHIVE"

tar czf "$ARCHIVE" -C "$PROJECT_ROOT" backend/storage/

SIZE=$(du -h "$ARCHIVE" | cut -f1)
echo "✓ Sauvegarde terminée ($SIZE)"

# Garder les 14 dernières sauvegardes
ls -t "$BACKUP_DIR"/devisflow-*.tar.gz 2>/dev/null | tail -n +15 | xargs -r rm
echo "✓ Anciennes sauvegardes nettoyées (gardées : 14 dernières)"
