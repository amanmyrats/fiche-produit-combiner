# fiche-produit-combiner
Combines seperate excel files(contains material cards) into one excel

Usage(Tested on Windows):
1 - Install virtual environment
    python -m venv venv

2 - Install requirements.txt
    pip install -r requirements.txt

3 - Convert to exe bundle. This will convert code into desktop application
    pyinstaller fp_menu.spec

4 - After converted into exe, copy folders and files below next to exe, exe file need them to run correctly.
Here are those folder and files:
Folder: app_files
File:   Fiche Produit Table.xlsm
File:   Fiche Produit Table.xlsx

5 - After copying dependent folders next to exe file, start application by clicking exe file.