#!/bin/bash

# Install Python
# winget install Python.Python.3.14

# Install python virtual envairoment
# python -m venv .venv

# Activate the virtual environment
source "./.venv/bin/activate"

# Install requirements.txt
./.venv/bin/pip install -r "./requirments.txt"

# Run PyInstaller on your script
./.venv/bin/python -m PyInstaller --windowed --add-data="./clock.jpeg:." --onefile "./gallop.py"
cp ./clock.jpeg ./dist
