#!/bin/sh
# Полный пересчет: симуляция -> Excel -> проверка Excel -> графики -> презентация PDF.
# Окружение (один раз): python3.11 -m venv .venv && .venv/bin/pip install numpy pandas scikit-learn openpyxl matplotlib formulas
set -e
cd "$(dirname "$0")"
PY=.venv/bin/python
$PY sim/simulate.py
$PY model/build_model.py
$PY model/check_model.py
$PY sim/charts.py
$PY deck/build_deck.py
