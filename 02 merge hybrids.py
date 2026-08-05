import os
import glob
import pandas as pd
from PyPDF2 import PdfWriter
from dotenv import load_dotenv

from files_flow_util import PDF_PDF_SHEET, ensure_files_flow

load_dotenv()


def short_path(path, file_manager=None):
    if not path:
        return path
    path = str(path)
    if file_manager and path.startswith(file_manager + os.sep):
        return os.path.join("File management", os.path.relpath(path, file_manager))
    parts = path.split(os.sep)
    if len(parts) > 3:
        return os.path.join(*parts[-3:])
    return path


def collect_hybrid_jobs(df, folder_prefix):
    """
    Each Excel column in PDF_PDF:
      row 0 -> final hybrid filename
      row 1+ -> 'subfolder, filename.pdf'
    """
    jobs = []
    for col in df.columns:
        serie = df[col].dropna()
        if serie.empty:
            continue

        final_name = str(serie.iloc[0]).strip()
        input_records = serie.iloc[1:]

        pieces = []
        missing = []
        bad_format = []

        for record in input_records:
            try:
                parts = [p.strip() for p in str(record).split(",")]
                if len(parts) != 2:
                    bad_format.append(str(record))
                    continue
                subfolder, filename = parts
                full_path = os.path.join(folder_prefix, subfolder, filename)
                if os.path.isfile(full_path):
                    pieces.append(full_path)
                else:
                    missing.append(full_path)
            except Exception as e:
                bad_format.append(f"{record} ({e})")

        jobs.append({
            "final_name": final_name,
            "pieces": pieces,
            "missing": missing,
            "bad_format": bad_format,
            "created": False,
            "error": None,
        })
    return jobs


def run_hybrids(jobs, output_folder):
    for job in jobs:
        if job["bad_format"]:
            job["error"] = "Bad format in Excel (expected 'Carpeta, archivo.pdf')"
            continue
        if job["missing"]:
            continue
        if not job["pieces"]:
            job["error"] = "No input pieces defined"
            continue

        dest_path = os.path.join(output_folder, job["final_name"])
        writer = PdfWriter()
        try:
            for path in job["pieces"]:
                writer.append(path)
            with open(dest_path, "wb") as f:
                writer.write(f)
            job["created"] = True
        except Exception as e:
            job["error"] = f"Merge failed: {e}"
        finally:
            writer.close()
    return jobs


def print_success(jobs):
    created = [j for j in jobs if j["created"]]
    print("\n" + "=" * 60)
    print("SUCCESS — hybrid files created")
    print("=" * 60)
    if not created:
        print("(none)")
        return
    for job in created:
        print(f"  {job['final_name']}  (piezas: {len(job['pieces'])})")
        for piece in job["pieces"]:
            print(f"         + {os.path.basename(piece)}")
    print(f"\nTotal success: {len(created)}")


def print_failure(jobs, file_manager):
    failed = [j for j in jobs if not j["created"]]
    print("\n" + "=" * 60)
    print("FAILURE — hybrid not created / missing pieces")
    print("=" * 60)
    if not failed:
        print("(none)")
        return

    for job in failed:
        print(f"\n  File: {job['final_name']}")
        print("  Missing to complete:")
        if job["missing"]:
            for path in job["missing"]:
                print(f"    - {short_path(path, file_manager)}")
        if job["bad_format"]:
            for item in job["bad_format"]:
                print(f"    - Bad Excel format: {item}")
        if job["error"] and not job["missing"] and not job["bad_format"]:
            print(f"    - {job['error']}")
        if not job["missing"] and not job["bad_format"] and not job["error"]:
            print("    - Unknown reason")

    print(f"\nTotal failure: {len(failed)}")


def print_report(jobs, file_manager):
    print_success(jobs)
    print_failure(jobs, file_manager)


def main():
    working_folder = os.getenv("working_folder")
    if not working_folder:
        print("❌ Set working_folder in .env")
        return

    excel_file = ensure_files_flow(working_folder)
    file_manager = os.path.join(working_folder, "File management")
    output_folder = os.path.join(file_manager, "hibridos PDF_PDF")

    try:
        # First row = output filename; remaining rows = piece paths
        df = pd.read_excel(excel_file, sheet_name=PDF_PDF_SHEET, header=None)
    except Exception as e:
        print(f"❌ Error al leer la hoja PDF_PDF: {e}")
        return

    os.makedirs(output_folder, exist_ok=True)

    non_pdf_files = [
        f for f in os.listdir(output_folder) if not f.lower().endswith(".pdf")
    ]
    if non_pdf_files:
        print(
            f"⚠️ Alerta: Se encontraron archivos no PDF en la carpeta de salida: "
            f"{non_pdf_files}"
        )
        print("Por favor, límpiala antes de continuar.")
        return

    for f in glob.glob(os.path.join(output_folder, "*.pdf")):
        os.remove(f)

    jobs = collect_hybrid_jobs(df, file_manager)
    jobs = run_hybrids(jobs, output_folder)
    print_report(jobs, file_manager)
    print(f"\n✨ Híbridos folder: {output_folder}")


if __name__ == "__main__":
    main()
