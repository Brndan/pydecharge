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

# def check_row(row):


def save_export_syndicats(export_sheet, output_file_path):
    wb = xlsx.Workbook()
    ws = wb.active
    header = ["Code organisation", "M. Mme", "Prénom", "Nom", "Heures décharges",
              "Minutes décharges", "Heures ORS", "Minutes ORS", "AIRE", "Corps", "RNE"]
    for i in range(len(header)):
        ws.cell(1,i+1).value = header[i]
    # On remplit le fichier ici avec le contenu de export_sheet
    for row in range(len(export_sheet)):
        export_sheet[row].insert(0, "S01")  # Code organisation → toujours S01
        export_sheet[row].insert(7, 0)  # Minutes ORS → toujours 0
        export_sheet[row].insert(-2, 2)  # Aire, toujours 2
        for cell in range(len(export_sheet[row])):
            ws.cell(row+2, cell+1).value = export_sheet[row][cell]
            ws.cell(row+2, cell+1).number_format = '@' # Appliquer le format 'Texte' évite une single quote avant les nombres
    wb.save(output_file_path)


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

    export_sheet = []

    for xlsx_file in all_xlsx:
        try:
            file = xlsx.load_workbook(filename=xlsx_file)
        except:
            sys.exit("Erreur à l’ouverture du fichier " + xlsx_file)
        sheet = file.active
        cell_range = sheet[args.begin:args.end]
        for row in range(len(cell_range)):
            export_row = []
            row_length = len(cell_range[row])
            empty_cells = 0
            for cell_coordinate in range(row_length):
                cell = cell_range[row][cell_coordinate].value
                if not cell:
                    cell = 0
                    empty_cells += 1
                export_row.append(cell)
            if not empty_cells == row_length:
                export_sheet.append(export_row)
    save_export_syndicats(export_sheet, output_file_path)


if __name__ == "__main__":
    main()
