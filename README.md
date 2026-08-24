# Risala — Arabic Correspondence Corpus Explorer

**Risala** (رسالة, Arabic for *letter*) is a digital humanities pipeline and web interface for the exploration of handwritten Arabic correspondence from the Larzac internment camp (1961–1962).

This repository accompanies the paper:

> Benabdallah, R. & Alrahabi, M. (2026). *Digitizing Confiscated Voices: A Computational Approach to Handwritten Arabic Letters from the Larzac Internment Camp (1961)*. Manuscript submitted for publication.

---

## Project Overview

The corpus consists of 1,623 handwritten Arabic letters intercepted by French authorities at the Larzac internment camp between August and September 1961. The pipeline automates metadata extraction and transcription using a multimodal large language model (Google Gemini), and provides a web-based interface for corpus exploration.

---

## Repository Structure

```
risala/
├── scripts/
│   ├── 0_extract_metadata_from_files_names.py   # Extract dates from folder names
│   ├── 1_extract_info_from_files_content.py     # LLM-based transcription & metadata extraction
│   ├── 2_excel_to_json_Interface.py             # Convert Excel BDD to JSON for the interface
│   └── 3_evaluation.py                          # CER/WER evaluation & metadata accuracy
├── interface/
│   └── risala.html                              # Standalone web interface
├── templates/
│   └── model_lettre.xml                         # TEI encoding template
├── .gitignore
├── LICENSE
└── README.md
```

---

## Pipeline

```
Scanned images (JPG)
        │
        ▼
0_extract_metadata_from_files_names.py
→ dates extracted from folder names (TSV)
        │
        ▼
1_extract_info_from_files_content.py
→ LLM transcription + metadata (output_llm.tsv)
        │
        ▼
2_excel_to_json_Interface.py
→ lettres_ar.json (for the web interface)
        │
        ▼
risala.html (web interface)
        │
        ▼
3_evaluation.py
→ CER/WER + metadata accuracy reports + figures
```

---

## Requirements

```bash
pip install google-genai openpyxl jiwer matplotlib pillow python-dotenv
```

---

## Configuration

Create a `.env` file in the project root:

```
GOOGLE_API_KEY=your_google_api_key_here
```

## Usage

### 1. Extract dates from folder names
```bash
python scripts/0_extract_metadata_from_files_names.py
```

### 2. Extract transcriptions and metadata via LLM
```bash
python scripts/1_extract_info_from_files_content.py
```
Configure `LIMIT_LETTERS`, `ORDER` (`sequential` or `random`), and `WHITELIST` at the top of the script.

### 3. Generate JSON for the interface
```bash
python scripts/2_excel_to_json_Interface.py
```

### 4. Launch the web interface
```bash
cd interface
python -m http.server 8888
# Then open http://localhost:8888/risala.html
```

Or use the provided `run.bat` on Windows.

### 5. Evaluate results
```bash
python scripts/3_evaluation.py
```
Set `MODE = "all"` | `"metadata"` | `"transcription"` at the top of the script.

---

## Web Interface Features

| Tab | Description |
|---|---|
| **Letters** | Filtered letter list with full-text detail panel |
| **Network** | Correspondence graph (sender/recipient nodes) |
| **Map** | Geographic map of sending/receiving places |
| **Timeline** | Chronological view (points, yearly/monthly histograms) |
| **Word Cloud** | Most frequent terms (Arabic/French filter) |
| **Clusters** | Automatic grouping by textual similarity (k-means / TF-IDF) |

---

## TEI Encoding

Each letter is encoded following the TEI P5 guidelines. See `templates/model_lettre.xml` for the encoding template, which covers:
- Correspondence metadata (`correspDesc`)
- Sender/recipient identification
- Automatic and manual transcriptions
- Inmate ID (`inmate_id`)

---

## Data Availability

The raw scanned images and the annotated Excel spreadsheet (`lettres.xlsx`) are not included in this repository due to archival sensitivity and data protection constraints. Researchers interested in accessing the corpus should contact the authors.

The JSON file (`lettres_ar.json`) used for the interface is available upon request.

---

## Citation

```bibtex
@unpublished{benabdallah2026risala,
  author = {Benabdallah, Radjaa and Alrahabi, Motasem},
  title  = {The {Risala} Corpus: Digitizing Arabic Letters from the {Larzac} Internment Camp (1961)},
  note   = {Manuscript submitted for publication},
  year   = {2026}
}
```

---

## License

Code: [MIT License](LICENSE)  
Data and corpus: [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)

---

## Acknowledgements

This work was carried out as part of an internship by **Radjaa Benabdallah**, supervised by **Motasem Alrahabi**, within the **ObTIC** research group, Sorbonne Université, Paris, France.
