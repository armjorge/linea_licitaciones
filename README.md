# Línea de licitaciones

Pipeline to turn a bookmarked Word/PDF consolidado into the final proposal folders (`LEGAL`, `TECNICO`, `ECONÓMICO`).

## Setup

```bash
cp .env.example .env          # set working_folder and final_proposal
cp files_flow.example.xlsx files_flow.xlsx   # or let scripts auto-create it
uv sync
# Linux Word→PDF needs LibreOffice:
# sudo apt-get install -y libreoffice-writer-nogui
```

`files_flow.xlsx` is gitignored (local working copy). If missing, scripts create it from `files_flow.example.xlsx` (or a blank template).

Folder layout under `working_folder`:

```text
DOC Templates/          consolidado .docx / .pdf
File management/
  output/               split PDFs from Word bookmarks
  fijos/                static PDFs you drop in manually
  bases/                other base files (xlsx/pdf)
  hibridos PDF_PDF/     merged static + output pieces
Propuesta_Final/        LEGAL | TECNICO | ECONÓMICO
files_flow.xlsx
```

## Three file types

| Tipo | Source | How it is produced |
|------|--------|--------------------|
| From Word bookmarks | `output/` | Split consolidado by heading bookmarks |
| Static | `fijos/` or `bases/` | You place the file manually |
| Hybrid (Word + static) | `hibridos PDF_PDF/` | `02 merge hybrids.py` merges pieces in order |

## `files_flow.xlsx` sheets

### `Parametrización` (proposal map)

Columns:

`id | Tipo | Requerimiento | Nota | File_name | Source | Source name | Move | letter_name | Word header`

| Column | Role |
|--------|------|
| `id` | Order within each `Move` (must be `1…N`) |
| `File_name` | Final name in `Propuesta_Final/{Move}/` (`.pdf` or `.xlsx`) |
| `Source` | Folder under `File management/` (`output`, `fijos`, `bases`, `hibridos PDF_PDF`) |
| `Source name` | Filename inside that Source folder |
| `Move` | Destination bucket: `LEGAL`, `TECNICO`, `ECONÓMICO` |
| `letter_name` | Optional alias for split output names |
| `Word header` | Bookmark title in the consolidado PDF (exact match) |

Example (Word-only row):

| id | Tipo | File_name | Source | Source name | Move | Word header |
|----|------|-----------|--------|-------------|------|-------------|
| 1 | Word | 01 ANEXO 3.pdf | output | 02_Anexo_3_acreditamiento_de.pdf | LEGAL | ANEXO 3 ACREDITAMIENTO DE PERSONALIDAD |

Example (hybrid row — no Word header; built by sheet `PDF_PDF`):

| id | Tipo | File_name | Source | Source name | Move |
|----|------|-----------|--------|-------------|------|
| 5 | Word-PDF | 05 Manifiesto.pdf | hibridos PDF_PDF | manifiesto_actuacion.pdf | LEGAL |

### `PDF_PDF` (hybrid recipes)

No header row. Each **column** is one hybrid:

- **Row 0:** output filename (e.g. `opinion_sat.pdf`)
- **Rows 1+:** pieces as `folder, filename.pdf` (merged top→bottom)

Example column:

```text
opinion_sat.pdf
output, 11_Cump_de_fiscal_emitida.pdf
fijos, opinion_sat.pdf
fijos, csf_sat.pdf
```

## Scripts (run in order)

```bash
uv run '01 Split PDF.py'      # Word→PDF bookmarks, list/check/split
uv run '02 merge hybrids.py'  # build hibridos from PDF_PDF
uv run '03 Mueve propuesta.py'# readiness report + wipe/copy to Propuesta_Final
uv run '04_Loading.py'        # copy File_name to clipboard by Move, one by one
```

### 01 Split PDF

1. Status / list bookmarks  
2. Generate bookmarked PDF from `.docx` (Word COM on Windows, LibreOffice on Linux)  
3. Check bookmarks vs `files_flow` `Word header`  
4. Split **every** bookmark → `File management/output/` (`01_….pdf`, `02_….pdf`, …)

### 02 merge hybrids

Reads `PDF_PDF`, merges pieces into `File management/hibridos PDF_PDF/`.  
Report: SUCCESS created / FAILURE missing pieces.

### 03 Mueve propuesta

1. Check readiness + gaps (by `Move` × `Source`, incomplete only)  
2. **Wipe** `final_proposal`, recreate Move folders, copy only ready files, verify no leftovers  

Ready = source file exists and `File_name` ends with `.pdf` / `.xlsx`.

### 04 Loading

Pick a `Move`, validate `id` is `1…N`, copy each `File_name` to the clipboard (Enter = next).

## Typical month-later checklist

1. Update consolidado in `DOC Templates/`  
2. Refresh `files_flow.xlsx` (`Word header`, `Source name`, `File_name`)  
3. Drop static PDFs into `File management/fijos/` (and `bases/` if needed)  
4. Run 01 → 02 → 03 → 04  
5. Upload from `Propuesta_Final/`
