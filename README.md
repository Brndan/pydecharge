# syndecharge

Compile les décharges individuelles des tableaux des syndicats.



## Utilisation

Usage : `syndecharge [--cts] [--begin A25] [--end J44] [-i dossier] [-o fichier]`
		
*syndecharge* est un programme qui compile les données de plages de fichiers
Excel contenant les déclarations de décharge des syndicats.
	

	--cts	crée un fichier Excel en sortie contenant les synthèses des syndicats :
	Syndicat | ETP attribué au syndicat | Mutualisation | "ETP disponibles | Consommé | Crédit d'Heures (CHS)
			Si --cts n'est pas renseigné, le défaut est une compilation des décharges.
		
	--begin	début de la plage de données. Par défaut, la valeur A25 est
	-b		attribuée à ce paramètre.
		
	--end	fin de la plage de données. Par défaut, la valeur J44 est
	-e		attribuée à ce paramètre.
		
	-input	dossier dans lequel se trouvent les fichiers à compiler. Si ce
	-i		paramètre est omis, le répertoire courant est utilisé.
		
	-output fichier en sortie. Par défaut, le fichier "export.xlsx" est
	-o		généré dans le répertoire courant.




*Exemples :*
	

	pydecharge.py --begin A25 --end J44 -i tableaux -o synthèse.xlsx

→ Génère un fichier synthèse.xlsx de la plage A25 à J44 de tous les fichiers
		situés dans le répertoire courant.
	

	pydecharge.py --cts --begin A74 --end A74

→ Génère un fichier export.xlsx dans le répertoire courant contenant la
synthèse de tous le temps utilisé par les syndicats à partir des fichiers
situés dans le répertoire courant.



## Dépendances

Le programme est codé en [Python](https://www.python.org/). 

La seule dépendance requise est le module `openpyxl`.

Pour installer les dépendances :

`pip install -r requirements.txt`