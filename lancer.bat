@echo off
cd /d "%~dp0"
python generer.py
start "" "http://localhost:8080/index.html"
python -m http.server 8080
