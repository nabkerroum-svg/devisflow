#!/usr/bin/env bash
#
# restore.sh — Restaure une sauvegarde DevisFlow
#
# Usage : ./scripts/restore.sh backups/devisflow-20260619-020000.tar.gz

set -e

if [ -z "$1" ]; then
  echo "Usage : $0 <archive.tar.gz>"
  echo ""
  echo "Sauvegardes disponibles :"
  ls -la "$(dirname "$0")/../backups/"devisflow-*.tar.gz 2>/dev/null || echo "  (aucune)"
  exit 1
fi

ARCHIVE="$1"
cd "$(dirname "$0")/.."
PROJECT_ROOT="$(pwd)"

if [ ! -f "$ARCHIVE" ]; then
  echo "❌ Archive introuvable : $ARCHIVE"
  exit 1
fi

echo "⚠️  Cette opération va REMPLACER toutes les données actuelles."
echo "    Source : $ARCHIVE"
read -p "Confirmer ? (oui/non) " CONFIRM
if [ "$CONFIRM" != "oui" ]; then
  echo "Annulé."
  exit 0
fi

# Sauvegarder l'existant par sécurité
if [ -d "backend/storage" ]; then
  BACKUP_EXISTING="backups/auto-before-restore-$(date +%Y%m%d-%H%M%S).tar.gz"
  mkdir -p backups
  tar czf "$BACKUP_EXISTING" -C "$PROJECT_ROOT" backend/storage/
  echo "✓ Sauvegarde de sécurité créée : $BACKUP_EXISTING"
fi

# Restaurer
rm -rf backend/storage
tar xzf "$ARCHIVE" -C "$PROJECT_ROOT"

echo "✓ Restauration terminée"
echo ""
echo "Redémarrer l'application maintenant :"
echo "  docker-compose down && docker-compose up -d"
echo "  ou : pkill uvicorn && uvicorn main:app  (en local)"
