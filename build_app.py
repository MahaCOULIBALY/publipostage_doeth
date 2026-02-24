#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de démarrage pour l'exécutable Publipostage DOETH.
Ce script sert de point d'entrée pour PyInstaller.
"""

import os
import sys
from pathlib import Path

# Ajout du répertoire du projet au path
app_path = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, app_path)

# Importation de l'interface graphique
from gui import main

if __name__ == "__main__":
    # Définir le répertoire de travail sur celui de l'exécutable
    if getattr(sys, 'frozen', False):
        os.chdir(os.path.dirname(sys.executable))
    main()
