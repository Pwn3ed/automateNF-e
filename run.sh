#!/usr/bin/env bash
set -euo pipefail

# Simple helper to create a venv, install requirements and run the script
PY=python3
VENV_DIR=".venv"

if [ ! -d "$VENV_DIR" ]; then
  echo "Creating virtual environment in $VENV_DIR..."
  $PY -m venv "$VENV_DIR"
fi

echo "Activating venv..."
source "$VENV_DIR/bin/activate"

echo "Installing requirements..."
pip install --upgrade pip
pip install -r requirements.txt

echo "Making geckodriver executable if present..."
if [ -f ./geckodriver ]; then
  chmod +x ./geckodriver || true
fi

echo "Running automate.py"
python automate.py
