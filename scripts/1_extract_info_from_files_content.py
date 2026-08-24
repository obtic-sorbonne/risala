"""
Ce script parcourt un corpus d'images de lettres (mon DD), regroupe les pages appartenant à une même lettre,
envoie les images à l'API Gemini pour extraire automatiquement les métadonnées et la transcription
en arabe, puis enregistre les résultats dans un fichier TSV (output_llm.tsv) destiné à être collé
manuellement dans la feuille "LLM" du classeur lettres.xlsx.

Structure de sortie (même colonnes que la feuille "Manuel") :
  #  folder_id  letter_id  sender_name  sender_place  date_sent  recipient_name
  recipient_place  date_received  inmate_id  transcription  notes

La colonne "notes" contient la liste des pages traitées (page_labels).
Une ligne par lettre (regroupement de toutes les pages ENV + CONT + DOC).
Reprise automatique : les lettres déjà présentes dans output_llm.tsv sont ignorées.
"""

import os
import time
import json
import random
from datetime import datetime
import csv
import re
from pathlib import Path
from dotenv import load_dotenv
from google import genai
from google.genai import types
from PIL import Image
import io

# ==============================================================================
# 1. CONFIGURATION
# ==============================================================================

load_dotenv()
client = genai.Client(api_key=os.getenv("GOOGLE_API_KEY"))

MODEL_NAME = "gemini-2.5-flash-lite"

# Chemin vers votre dossier racine d'images
BASE_DIR = Path(r"C:\Users\sorbonne\Documents\Workspace\! Corpus\arabicCorpus\Rajaa")

# Fichier de sortie TSV (à coller dans la feuille "LLM" du classeur lettres.xlsx)
OUTPUT_TSV = Path(r"C:\Users\sorbonne\Documents\Workspace\Risala\output\output_llm.tsv")

# Limite de session (None = traiter tout le corpus)
LIMIT_LETTERS = 50

# Ordre de traitement : "sequential" ou "random"
ORDER = "random"

# Liste blanche de lettres à traiter ([] = tout le corpus)
# Format : [("folder_id", "letter_id"), ...]
WHITELIST = []
#     ("1_03081961", "L1"),
#     ("1_03081961", "L10"),
#     ("1_03081961", "L11"),
#     ("1_03081961", "L14"),
#     ("1_03081961", "L15"),
#     ("1_03081961", "L2"),
#     ("1_03081961", "L20"),
#     ("1_03081961", "L21"),
#     ("1_03081961", "L22"),
#     ("1_03081961", "L24"),
#     ("1_03081961", "L25"),
#     ("1_03081961", "L29"),
#     ("1_03081961", "L3"),
#     ("1_03081961", "L31"),
#     ("1_03081961", "L33"),
#     ("1_03081961", "L37"),
#     ("1_03081961", "L6"),
#     ("1_03081961", "L9"),
# ]

# Résolution maximale des images (réduit le coût ~9x)
IMAGE_MAX_SIZE = 1024

# ==============================================================================
# 2. PROMPT SYSTÈME
# ==============================================================================
SYSTEM_INSTRUCTIONS = """
Tu es un expert en paléographie et en archivage de documents historiques en langue arabe.
Analyse l'ensemble des images fournies, qui constituent une seule et même lettre.

Les images peuvent être de plusieurs types, chacun pouvant avoir plusieurs pages numérotées :
- ENV (1 ou plusieurs pages) : enveloppe recto/verso → source principale pour sender_name, sender_place, recipient_name, recipient_place, date_sent, date_received, inmate_id
- CONT (1 ou plusieurs pages) : pages manuscrites du contenu → source pour la transcription et éventuellement date_sent, sender_place si absents de l'enveloppe
- DOC (1 ou plusieurs pages) : documents administratifs joints → source complémentaire pour inmate_id uniquement

Règles de priorité pour les métadonnées :
- sender_name    : ENV en priorité (nom écrit sur l'enveloppe), sinon signature dans CONT
- sender_place   : cachet postal ENV en priorité, sinon en-tête CONT
- date_sent      : cachet postal ENV en priorité, sinon début CONT
- recipient_name : face de l'ENV uniquement
- recipient_place: face de l'ENV uniquement
- date_received  : cachet de réception au dos de l'ENV uniquement
- inmate_id      : tampon administratif ENV ou DOC uniquement

Normalisation de la casse : première lettre en majuscule, reste en minuscules pour chaque mot
(ex: "Larzac", "Mohamed Ali", "France"). Applique cette règle à tous les champs sauf "transcription".

Si une information est absente, laisse la chaîne vide "".
IMPORTANT : Réponds UNIQUEMENT avec le JSON brut, sans balises markdown, sans ```json.

Schéma JSON attendu :
{
    "sender_name": "Nom de l'expéditeur",
    "sender_place": "Lieu d'expédition, format JJ/MM/AAAA",
    "date_sent": "Date d'envoi, format JJ/MM/AAAA",
    "recipient_name": "Nom du destinataire",
    "recipient_place": "Lieu de réception",
    "date_received": "Date de réception, format JJ/MM/AAAA",
    "inmate_id": "Numéro de matricule",
    "transcription": "Transcription intégrale, fidèle et mot à mot de TOUT le texte arabe manuscrit visible sur TOUTES les pages CONT, dans l'ordre des pages."
}
"""

# ==============================================================================
# 3. FONCTIONS OUTILS
# ==============================================================================

def encode_image(image_path, max_size=IMAGE_MAX_SIZE):
    """Redimensionne et convertit une image en bytes pour réduire le coût API."""
    img = Image.open(image_path)
    img.thumbnail((max_size, max_size))
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=85)
    return buffer.getvalue()


def get_existing_letters():
    """Lit le TSV existant et retourne les couples (folder_id, letter_id) déjà traités."""
    if not OUTPUT_TSV.exists():
        return set()
    processed = set()
    with open(OUTPUT_TSV, "r", encoding="utf-8") as f:
        reader = csv.reader(f, delimiter="\t")
        next(reader, None)  # skip header
        for row in reader:
            if len(row) >= 3:
                processed.add((row[1], row[2]))  # folder_id, letter_id
    return processed


def safe(val):
    """Retourne une chaîne propre depuis une valeur JSON potentiellement None."""
    return str(val or "").strip()


# ==============================================================================
# 4. INITIALISATION DU FICHIER DE SORTIE
# ==============================================================================
if not OUTPUT_TSV.exists():
    with open(OUTPUT_TSV, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter="\t")
        writer.writerow([
            "#", "folder_id", "letter_id",
            "sender_name", "sender_place", "date_sent",
            "recipient_name", "recipient_place", "date_received",
            "inmate_id", "transcription", "notes"
        ])

processed_letters = get_existing_letters()
row_counter = len(processed_letters)

# ==============================================================================
# 5. EXPLORATION DE L'ARBORESCENCE ET REGROUPEMENT DES IMAGES
# ==============================================================================
print("Analyse de l'arborescence du corpus en cours...")

corpus_groups = {}

for folder in sorted(BASE_DIR.iterdir()):
    if folder.is_dir() and not folder.name.startswith("!"):
        folder_id = re.sub(r'^ENV', '', folder.name)
        corpus_groups[folder_id] = {}
        for img_path in sorted(folder.glob("*.jpg")):
            match = re.match(r"^(L\d+)_", img_path.name)
            if match:
                letter_prefix = match.group(1)
                if letter_prefix not in corpus_groups[folder_id]:
                    corpus_groups[folder_id][letter_prefix] = []
                corpus_groups[folder_id][letter_prefix].append(img_path)

total_letters = sum(
    len(letters)
    for letters in corpus_groups.values()
    for _ in [letters]
)
total_letters = sum(len(v) for v in corpus_groups.values())
processed_in_this_run = 0
session_start = time.time()
times_per_letter = []

print(f"Total lettres identifiées  : {total_letters}")
print(f"Déjà dans output_llm.tsv   : {len(processed_letters)}")
if LIMIT_LETTERS:
    print(f"Limite de session          : {LIMIT_LETTERS} lettres")

# ==============================================================================
# 6. BOUCLE PRINCIPALE
# ==============================================================================

# Construction de la liste de toutes les lettres à traiter
all_tasks = [
    (folder_id, letter_id, img_list)
    for folder_id, letters in corpus_groups.items()
    for letter_id, img_list in letters.items()
]

if ORDER == "random":
    random.shuffle(all_tasks)
    print(f"Ordre : aléatoire (seed non fixée)")
else:
    print(f"Ordre : successif")

for folder_id, letter_id, img_list in all_tasks:

    if LIMIT_LETTERS and processed_in_this_run >= LIMIT_LETTERS:
        break

    # Filtre liste blanche
    if WHITELIST and (folder_id, letter_id) not in WHITELIST:
        continue

    # Reprise automatique
    if (folder_id, letter_id) in processed_letters:
        continue

    row_counter += 1
    processed_in_this_run += 1
    session_target = LIMIT_LETTERS if LIMIT_LETTERS else total_letters

    print(f"[{row_counter}/{total_letters}] "
          f"Session:[{processed_in_this_run}/{session_target}] "
          f"Dossier: {folder_id} | Lettre: {letter_id} "
          f"({len(img_list)} images)...")

    letter_start = time.time()

    # Liste des noms de pages (page_labels)
    page_labels = ", ".join(p.name for p in img_list)

    # Construction du payload : prompt + images redimensionnées
    parts = [SYSTEM_INSTRUCTIONS]
    for img_path in img_list:
        try:
            img_bytes = encode_image(img_path)
            parts.append(types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg"))
        except Exception as e:
            print(f"   ⚠️ Erreur image {img_path.name}: {e}")

    # Envoi à Gemini
    try:
        response = client.models.generate_content(
            model=MODEL_NAME,
            contents=parts,
            config=types.GenerateContentConfig(
                temperature=0.1,
                response_mime_type="application/json"
            )
        )

        if not response or not response.text:
            print(f"   ⚠️ Réponse vide pour {letter_id}, ignorée.")
            continue

        raw = response.text.strip()
        raw = re.sub(r'^```json\s*', '', raw)
        raw = re.sub(r'\s*```$', '', raw)

        try:
            result = json.loads(raw)
        except json.JSONDecodeError:
            def extract_field(text, key):
                m = re.search(rf'"{key}"\s*:\s*"(.*?)"(?=\s*[,}}])', text, re.DOTALL)
                return m.group(1) if m else ""
            result = {k: extract_field(raw, k) for k in [
                "sender_name", "sender_place", "date_sent",
                "recipient_name", "recipient_place", "date_received",
                "inmate_id", "transcription"
            ]}

        # Nettoyage de la transcription
        transcription = safe(result.get("transcription")).replace("\n", " ").replace("\r", " ").replace("\t", " ")

        # Écriture dans le TSV
        with open(OUTPUT_TSV, "a", encoding="utf-8", newline="") as f:
            writer = csv.writer(f, delimiter="\t")
            writer.writerow([
                row_counter,
                folder_id,
                letter_id,
                safe(result.get("sender_name")),
                safe(result.get("sender_place")),
                safe(result.get("date_sent")),
                safe(result.get("recipient_name")),
                safe(result.get("recipient_place")),
                safe(result.get("date_received")),
                safe(result.get("inmate_id")),
                transcription,
                f"pages: {page_labels}"
            ])

        elapsed_letter = time.time() - letter_start
        times_per_letter.append(elapsed_letter)
        print(f"   ⏱ {elapsed_letter:.1f}s")
        time.sleep(1.0)

    except Exception as e:
        print(f"❌ Échec {letter_id} / {folder_id}: {e}")
        if '503' in str(e):
            print("   ⏳ Serveur surchargé, attente 30s...")
            time.sleep(30)
        else:
            time.sleep(5)

session_elapsed = time.time() - session_start
avg_time = sum(times_per_letter) / len(times_per_letter) if times_per_letter else 0

print(f"\n🎉 Session achevée. {processed_in_this_run} lettres traitées.")
print(f"⏱  Durée totale        : {session_elapsed/60:.1f} min ({session_elapsed:.0f}s)")
print(f"⏱  Temps moyen/lettre  : {avg_time:.1f}s")
print(f"📋 Requêtes API        : {processed_in_this_run} (une par lettre)")
print(f"💾 Fichier TSV         : {OUTPUT_TSV}")
print("→ Copiez le contenu dans la feuille 'LLM' de lettres.xlsx")
