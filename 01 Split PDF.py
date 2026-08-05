import os
import shutil
import subprocess
import tempfile
from PyPDF2 import PdfReader, PdfWriter
#import xlsxwriter
import sys
from dotenv import load_dotenv
load_dotenv()
import unicodedata
import re
import pandas as pd
import glob

from files_flow_util import ensure_files_flow

if sys.platform == "win32":
    import win32com.client as win32
else:
    win32 = None

working_folder = os.getenv("working_folder")
DEFAULT_DOC_NAME = "CONSOLIDADA_2027"
FILES_FLOW_SHEET = "Parametrización"


def templates_dir():
    return os.path.join(working_folder, "DOC Templates")


def resolve_paths(doc_name=DEFAULT_DOC_NAME):
    base = templates_dir()
    return {
        "doc_name": doc_name,
        "templates_dir": base,
        "pdf_path": os.path.join(base, f"{doc_name}.pdf"),
        "word_path": os.path.join(base, f"{doc_name}.docx"),
        "headers_xlsx": os.path.join(base, "headers_files.xlsx"),
        "files_flow": ensure_files_flow(working_folder),
        "output_folder": os.path.join(working_folder, "File management", "output"),
    }


def generate_bookmark_pdf(word_path):
    """
    Convert a Word .docx to PDF with heading bookmarks preserved.
    Windows: Microsoft Word COM (CreateBookmarks from headings).
    Linux: LibreOffice headless writer_pdf_Export with ExportBookmarks.
    Returns the PDF path on success, None on failure.
    """
    if not os.path.isfile(word_path):
        print(f"Error: Word file not found: {word_path}")
        return None

    pdf_path = os.path.splitext(word_path)[0] + ".pdf"
    print(f"Generating bookmark PDF from: {word_path}")

    if sys.platform == "win32":
        ok = _generate_bookmark_pdf_windows(word_path, pdf_path)
    else:
        ok = _generate_bookmark_pdf_linux(word_path, pdf_path)

    if ok and os.path.isfile(pdf_path):
        print(f"PDF successfully created: {os.path.abspath(pdf_path)}")
        return pdf_path

    print(f"Error: Failed to generate PDF: {pdf_path}")
    return None


def _generate_bookmark_pdf_windows(word_path, pdf_path):
    if win32 is None:
        print("Error: pywin32 / win32com is not available.")
        return False

    word = None
    doc = None
    try:
        word = win32.gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(word_path))
        doc.ExportAsFixedFormat(
            OutputFileName=os.path.abspath(pdf_path),
            ExportFormat=17,  # PDF
            OpenAfterExport=False,
            OptimizeFor=0,
            CreateBookmarks=1,  # bookmarks from headings
        )
        return True
    except Exception as e:
        print(f"Error: Word COM export failed: {e}")
        return False
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception as close_error:
                print(f"Warning: Could not close Word document: {close_error}")
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def _find_libreoffice():
    for candidate in ("soffice", "libreoffice"):
        path = shutil.which(candidate)
        if path:
            return path
    return None


def _generate_bookmark_pdf_linux(word_path, pdf_path):
    soffice = _find_libreoffice()
    if not soffice:
        print(
            "Error: LibreOffice not found. Install it with:\n"
            "  sudo apt-get install -y libreoffice-writer-nogui"
        )
        return False

    outdir = os.path.dirname(os.path.abspath(pdf_path)) or "."
    expected_pdf = os.path.join(
        outdir, os.path.splitext(os.path.basename(word_path))[0] + ".pdf"
    )

    # Unique LO user profile avoids lock conflicts with a running GUI instance.
    with tempfile.TemporaryDirectory(prefix="lo_profile_") as profile_dir:
        filter_opts = (
            'pdf:writer_pdf_Export:'
            '{"ExportBookmarks":{"type":"boolean","value":"true"}}'
        )
        cmd = [
            soffice,
            "--headless",
            f"-env:UserInstallation=file://{profile_dir}",
            "--convert-to",
            filter_opts,
            "--outdir",
            outdir,
            os.path.abspath(word_path),
        ]
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=180,
                check=False,
            )
        except subprocess.TimeoutExpired:
            print("Error: LibreOffice conversion timed out.")
            return False
        except Exception as e:
            print(f"Error: LibreOffice conversion failed to start: {e}")
            return False

        if result.returncode != 0:
            print(f"Error: LibreOffice exited with code {result.returncode}")
            if result.stderr:
                print(result.stderr.strip())
            if result.stdout:
                print(result.stdout.strip())
            return False

    if not os.path.isfile(expected_pdf):
        print(f"Error: LibreOffice did not produce: {expected_pdf}")
        if result.stdout:
            print(result.stdout.strip())
        return False

    if os.path.abspath(expected_pdf) != os.path.abspath(pdf_path):
        shutil.move(expected_pdf, pdf_path)

    return True


def flatten_bookmarks(outline):
    """Flatten nested PDF outlines (Heading 1/2/3 trees) into a list of destinations."""
    flat = []
    for item in outline or []:
        if isinstance(item, list):
            flat.extend(flatten_bookmarks(item))
        elif isinstance(item, dict):
            flat.append(item)
    return flat


def clean_filename(text, index):
    # 1. Quitar acentos y normalizar
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    text = text.upper()

    # 2. Abreviaturas clave para licitaciones
    replacements = {
        'MANIFESTACION': 'MANIF',
        'ARTICULOS': 'ART',
        'OBLIGACIONES FISCALES': 'FISCAL',
        'CUMPLIMIENTO': 'CUMP',
        'PERSONALIDAD JURIDICA': 'PERS_JUR',
        'CONFLICTO DE INTERESES': 'NO_CONF_INT'
    }
    for word, rep in replacements.items():
        text = text.replace(word, rep)

    # 3. Limpiar caracteres no permitidos (deja letras, números y espacios)
    text = re.sub(r'[^A-Z0-9\s]', '', text)

    # 4. Acortar: tomamos las primeras 4 palabras y agregamos el índice para orden
    words = text.split()
    short_name = "_".join(words[:4])

    # Retornamos formato: 01_Nombre_Corto.pdf
    return f"{str(index).zfill(2)}_{short_name.capitalize()}.pdf"


def load_pdf_bookmarks(pdf_path):
    """Return (PdfReader, flat bookmark destinations) or (None, [])."""
    if not os.path.isfile(pdf_path):
        print(f"❌ PDF not found: {pdf_path}")
        return None, []
    reader = PdfReader(pdf_path)
    bookmarks = flatten_bookmarks(reader.outline)
    return reader, bookmarks


def physical_bookmark_map(pdf_reader, bookmarks):
    """Build {title: start_page} from flattened bookmarks."""
    physical = {}
    for item in flatten_bookmarks(bookmarks):
        title = item.get("/Title")
        if not title:
            continue
        physical[title] = pdf_reader.get_destination_page_number(item)
    return physical


def load_files_flow(files_flow_path):
    if not os.path.isfile(files_flow_path):
        print(f"❌ No se encontró el archivo: {files_flow_path}")
        return None
    return pd.read_excel(files_flow_path, sheet_name=FILES_FLOW_SHEET)


def export_headers_xlsx(bookmarks, output_path):
    extracted = []
    for i, item in enumerate(bookmarks):
        title = item.get("/Title")
        if not title:
            continue
        extracted.append({
            "word_headers": title,
            "sanitized_name": clean_filename(title, i + 1),
        })
    if not extracted:
        print("⚠️ No se encontraron bookmarks para exportar.")
        return False
    pd.DataFrame(extracted).to_excel(output_path, index=False)
    print(f"✅ Excel generado con {len(extracted)} registros: {output_path}")
    return True


def show_status(paths):
    print("\n--- Document status ---")
    print(f"working_folder : {working_folder}")
    print(f"DOC name       : {paths['doc_name']}")
    print(f"Word           : {paths['word_path']}  [{'OK' if os.path.isfile(paths['word_path']) else 'MISSING'}]")
    print(f"PDF            : {paths['pdf_path']}  [{'OK' if os.path.isfile(paths['pdf_path']) else 'MISSING'}]")
    print(f"files_flow     : {paths['files_flow']}  [{'OK' if os.path.isfile(paths['files_flow']) else 'MISSING'}]")

    reader, bookmarks = load_pdf_bookmarks(paths["pdf_path"])
    if reader is None:
        return

    titles = [b.get("/Title") for b in bookmarks if b.get("/Title")]
    print(f"Pages          : {len(reader.pages)}")
    print(f"Bookmarks      : {len(titles)}")
    for title in titles:
        print(f"  • {title}")

    try:
        export_headers_xlsx(bookmarks, paths["headers_xlsx"])
    except Exception as e:
        print(f"❌ Error al salvar Excel: {e}")


def action_generate_pdf(paths):
    if not os.path.isfile(paths["word_path"]):
        print(f"❌ Word file missing: {paths['word_path']}")
        return
    if os.path.isfile(paths["pdf_path"]):
        answer = input("PDF already exists. Overwrite? [y/N]: ").strip().lower()
        if answer not in ("y", "yes", "s", "si", "sí"):
            print("Cancelled.")
            return
        os.remove(paths["pdf_path"])
    generate_bookmark_pdf(paths["word_path"])


def check_vs_files_flow(paths):
    """Compare PDF bookmarks against files_flow.xlsx and summarize gaps."""
    reader, bookmarks = load_pdf_bookmarks(paths["pdf_path"])
    if reader is None:
        return

    df_split = load_files_flow(paths["files_flow"])
    if df_split is None:
        return

    physical = physical_bookmark_map(reader, bookmarks)
    pdf_titles = set(physical.keys())

    mask_missing = df_split["Word header"].isna() & df_split["letter_name"].notna()
    mask_inverse = df_split["letter_name"].isna() & df_split["Word header"].notna()
    incomplete = df_split[mask_missing | mask_inverse]

    excel_headers = df_split["Word header"].dropna().tolist()
    excel_set = set(excel_headers)
    missing_in_pdf = [h for h in excel_headers if h not in physical]
    extra_in_pdf = sorted(pdf_titles - excel_set)
    matched = sorted(pdf_titles & excel_set)

    print("\n" + "=" * 50)
    print("CHECK: PDF bookmarks vs files_flow.xlsx")
    print("=" * 50)
    print(f"PDF bookmarks     : {len(pdf_titles)}")
    print(f"Excel Word headers: {len(excel_headers)}")
    print(f"Matched           : {len(matched)}")
    print(f"In Excel, not PDF : {len(missing_in_pdf)}")
    print(f"In PDF, not Excel : {len(extra_in_pdf)}")
    print(f"Incomplete Excel  : {len(incomplete)}")

    if len(incomplete):
        print("\n⚠️ Incomplete Excel rows (letter_name / Word header mismatch):")
        cols = [c for c in ["id", "Source", "letter_name", "Word header"] if c in incomplete.columns]
        print(incomplete[cols].to_string(index=False))

    if missing_in_pdf:
        print("\n🛑 In files_flow but MISSING in PDF bookmarks:")
        for h in missing_in_pdf:
            print(f"   - {h}")
    else:
        print("\n✅ Every Excel Word header exists in the PDF.")

    if extra_in_pdf:
        print("\n📌 In PDF bookmarks but NOT mapped in files_flow:")
        for h in extra_in_pdf:
            print(f"   - {h}")
    else:
        print("\n✅ Every PDF bookmark is mapped in files_flow.")

    if matched:
        print(f"\n✅ Matched headers ({len(matched)}):")
        for h in matched:
            print(f"   - {h}")

    print("=" * 50)


def bookmark_split_plan(pdf_reader, bookmarks):
    """
    One output file per bookmark, in document order.
    Filenames match headers_files.xlsx (sanitized_name with index prefix).
    """
    plan = []
    for item in flatten_bookmarks(bookmarks):
        title = item.get("/Title")
        if not title:
            continue
        start_page = pdf_reader.get_destination_page_number(item)
        plan.append({
            "title": title,
            "start": start_page,
            "filename": clean_filename(title, len(plan) + 1),
        })

    for i, entry in enumerate(plan):
        if i + 1 < len(plan):
            entry["end"] = plan[i + 1]["start"]
        else:
            entry["end"] = len(pdf_reader.pages)
    return plan


def split_pdf_files(paths):
    """Split every PDF bookmark into its own file (independent of files_flow)."""
    reader, bookmarks = load_pdf_bookmarks(paths["pdf_path"])
    if reader is None:
        return

    plan = bookmark_split_plan(reader, bookmarks)
    if not plan:
        print("⚠️ No bookmarks found in PDF. Nothing to split.")
        return

    try:
        export_headers_xlsx(bookmarks, paths["headers_xlsx"])
    except Exception as e:
        print(f"⚠️ Could not refresh headers_files.xlsx: {e}")

    print(f"\n✅ Splitting {len(plan)} bookmark(s) into individual PDFs...")
    split_pdf_by_bookmarks(reader, plan, paths["output_folder"])


def split_pdf_by_bookmarks(pdf_reader, plan, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"📁 Carpeta de salida creada: {output_folder}")
    else:
        print("✅ Carpeta de salida detectada.")

    files_to_delete = glob.glob(os.path.join(output_folder, "*.pdf"))
    for f in files_to_delete:
        os.remove(f)
    if files_to_delete:
        print(f"🗑️ Se eliminaron {len(files_to_delete)} archivos PDF previos.")

    for entry in plan:
        start_page = entry["start"]
        end_page = entry["end"]
        output_filename = entry["filename"]

        if end_page <= start_page:
            print(
                f"⚠️ Skip empty range for '{entry['title']}' "
                f"(start page {start_page + 1})"
            )
            continue

        writer = PdfWriter()
        for page_idx in range(start_page, end_page):
            writer.add_page(pdf_reader.pages[page_idx])

        dest_path = os.path.join(output_folder, output_filename)
        with open(dest_path, "wb") as f:
            writer.write(f)

        print(
            f"📄 Generado: {output_filename} "
            f"(Págs {start_page + 1}-{end_page}) — {entry['title']}"
        )

    print(f"\n✨ Proceso terminado. Archivos listos en: {output_folder}")


def print_menu(paths):
    pdf_ok = "OK" if os.path.isfile(paths["pdf_path"]) else "MISSING"
    word_ok = "OK" if os.path.isfile(paths["word_path"]) else "MISSING"
    print("\n" + "=" * 50)
    print("SPLIT PDF MENU")
    print("=" * 50)
    print(f"DOC : {paths['doc_name']}")
    print(f"Word: [{word_ok}] {os.path.basename(paths['word_path'])}")
    print(f"PDF : [{pdf_ok}] {os.path.basename(paths['pdf_path'])}")
    print("-" * 50)
    print("1. Show status / list bookmarks")
    print("2. Generate bookmarked PDF from Word")
    print("3. Check PDF vs files_flow.xlsx")
    print("4. Split PDF by bookmarks (one file per bookmark)")
    print("5. Change document name")
    print("0. Exit")
    print("=" * 50)


def main():
    doc_name = DEFAULT_DOC_NAME
    paths = resolve_paths(doc_name)

    while True:
        print_menu(paths)
        choice = input("Choose an option: ").strip()

        if choice == "1":
            show_status(paths)
        elif choice == "2":
            action_generate_pdf(paths)
        elif choice == "3":
            check_vs_files_flow(paths)
        elif choice == "4":
            split_pdf_files(paths)
        elif choice == "5":
            new_name = input(
                f"Document name without extension [{paths['doc_name']}]: "
            ).strip()
            if new_name:
                doc_name = new_name
                paths = resolve_paths(doc_name)
                print(f"Using document: {doc_name}")
        elif choice == "0":
            print("Bye.")
            break
        else:
            print("Invalid choice. Please select 0-5.")
            continue

        input("\nPress Enter to return to menu...")


if __name__ == "__main__":
    main()
