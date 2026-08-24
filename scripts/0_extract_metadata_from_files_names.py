"""
Ce script parcourt les dossiers nommés selon le format `ENV..._JJMMAAAA`, extrait la date du nom du dossier,
la convertit au format `AAAA-MM-JJ`, puis enregistre dans un fichier TSV (`metadata.csv`) une ligne par image
contenant l'identifiant du dossier (`folder_id`), l'identifiant de la lettre (nom du fichier sans extension,
`letter_id`) et la date d'envoi (`date_sent`):

folder_id    letter_id    date_sent
ENV1_03081961    L1_CONT_1.jpg    1961-08-03

ces données seront collées manuellement dans le document lettres algériennes (drive) dans la feuille "Manuel"
elles seront ensuite complétées par d'autres métadonnées (voir drive)

"""
 
from pathlib import Path
import csv
import re
from datetime import datetime

ROOT_DIR = Path(r"C:\Users\sorbonne\Documents\Workspace\! Corpus\arabicCorpus\Rajaa")
OUTPUT_FILE = Path(r"C:\Users\sorbonne\Documents\Workspace\Risala\output\metadata.csv")

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp"}

folder_pattern = re.compile(r"^ENV\d+_(\d{8})$")

with OUTPUT_FILE.open("w", encoding="utf-8", newline="") as csvfile:
    writer = csv.writer(csvfile, delimiter="\t")
    writer.writerow(["folder_id", "letter_id", "date_sent"])

    for folder in sorted(ROOT_DIR.iterdir()):
        if not folder.is_dir():
            continue

        match = folder_pattern.match(folder.name)
        if not match:
            continue

        raw_date = match.group(1)
        date_sent = datetime.strptime(raw_date, "%d%m%Y").strftime("%Y-%m-%d")

        for image in sorted(folder.iterdir()):
            if image.is_file() and image.suffix.lower() in IMAGE_EXTENSIONS:
                folder_id = folder.name
                letter_id = image.stem
                writer.writerow([folder_id, letter_id, date_sent])