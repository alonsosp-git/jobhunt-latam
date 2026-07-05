#!/usr/bin/env bash
cd "$(dirname "$0")"
echo "Installing/updating dependencies (first run only)..."
python3 -m pip install -r requirements.txt
echo
echo "Starting JobHunt LatAm... open http://127.0.0.1:5000 in your browser."
python3 app.py
