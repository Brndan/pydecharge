#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Tous les fichiers, en entrée comme en sortie, sont des XLSX.


import sys
import os
from pathlib import Path
import shutil
import argparse

import openpyxl as xlsx


def main():
    parser = argparse.ArgumentParser(
        description="Produire à partir d’un modèle les fichiers de décharge pour tous les syndicats.")
    parser.add_argument(
        "--cts",
        help="Synthèse de la ligne des CTS des syndicats",
        required=False)
    parser.add_argument(
        "--begin",
        "-b",
        action="store",
        help="Début de la plage de données",
        required=False
    )
    parser.add_argument(
        "--end",
        "-e",
        action="store",
        help="Fin de la plage de données",
        required=False
    )
    parser.add_argument(
        "-i",
        "--input",
        action="store",
        help="dossier dans lequel se trouvent les fichiers à compiler. Si ce paramètre est omis, le répertoire courant est utilisé.",
        required=False
    )
    parser.add_argument(
        "-o",
        "--output",
        action="store",
        help="fichier en sortie. Par défaut, le fichier export.xlsx est généré dans le répertoire courant.",
        required=False
    )
    parser.add_argument(
        "--csv",
        help="Synthèse de la ligne des CTS des syndicats",
        required=False)
    args = parser.parse_args()

    if args.input:
        source_folder = Path(args.input)
    else:
        source_folder = os.getcwd()

    if args.output:
        output_file_path = Path(args.output)
    else:
        output_file_path = os.path.join(os.getcwd(), "export.xlsx")
    print(output_file_path)

    if not args.begin:
        args.begin = "B25"
    if not args.end:
        args.end = "J44"
    

if __name__ == "__main__":
    main()
