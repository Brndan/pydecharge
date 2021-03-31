#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Tous les fichiers, en entrée comme en sortie, sont des XLSX.


import sys
import os
from pathlib import Path
import shutil
import argparse

import glob

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

    if not args.begin:
        args.begin = "B25"
    if not args.end:
        args.end = "J44"

    if not os.path.exists(source_folder):
        sys.exit("La source spécifiée n’existe pas.")

    # Récupère une liste de tous les xlsx dans le dossier source
    all_xlsx = glob.glob(os.path.join(source_folder, "*.xlsx"))

    #

    for xlsx_file in all_xlsx:
        try:
            file = xlsx.load_workbook(filename=xlsx_file)
        except:
            sys.exit("Erreur à l’ouverture du fichier " + xlsx_file)
        sheet = file.active
        cell_range = sheet[args.begin:args.end]
        print(cell_range[0][0].value)

if __name__ == "__main__":
    main()
