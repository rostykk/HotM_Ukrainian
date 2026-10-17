#!/bin/sh
# Видалення українізатора Heart of the Machine (Linux / Steam Deck)
HOTM_UK_MODE=uninstall exec sh "$(cd "$(dirname "$0")" && pwd)/install.sh" "$@"
