#!/bin/bash
# Inicia o servidor do Gerador de Certificados
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PORT="${PORT:-8080}"

# Verifica Python 3
if ! command -v python3 &>/dev/null; then
  echo "ERRO: Python 3 não encontrado."
  exit 1
fi

# Verifica Microsoft PowerPoint via AppleScript
PP_VERSION=$(osascript -e 'tell application "Microsoft PowerPoint" to version' 2>/dev/null || true)
if [ -z "$PP_VERSION" ]; then
  echo "AVISO: Microsoft PowerPoint não encontrado. A conversão para PDF não funcionará."
else
  echo "PowerPoint $PP_VERSION detectado."
fi

echo ""
echo "Abrindo http://localhost:$PORT ..."
open "http://localhost:$PORT" 2>/dev/null &

cd "$SCRIPT_DIR"
PORT=$PORT python3 server.py
