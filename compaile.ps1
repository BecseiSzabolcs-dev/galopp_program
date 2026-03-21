# Install Python 
# winget install Python.Python.3.14

# Install python virtual envairoment
# python -m venv .venv

# Install requirements.txt
& ".\.venv\Scripts\pip.exe" install -r "./requirements.txt"

# Activate the virtual environment
& ".\.venv\Scripts\Activate.ps1"

# Run PyInstaller on your script
& ".\.venv\Scripts\python.exe" -m PyInstaller --windowed --onefile ".\ugeto.py"