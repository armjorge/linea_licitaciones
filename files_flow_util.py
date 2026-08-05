"""Shared helpers for files_flow.xlsx bootstrap and paths."""

from __future__ import annotations

import os
import shutil

import pandas as pd

PARAM_HEADERS = [
    "id",
    "Tipo",
    "Requerimiento",
    "Nota",
    "File_name",
    "Source",
    "Source name",
    "Move",
    "letter_name",
    "Word header",
]

PARAM_SHEET = "Parametrización"
PDF_PDF_SHEET = "PDF_PDF"


def files_flow_path(working_folder: str) -> str:
    return os.path.join(working_folder, "files_flow.xlsx")


def example_files_flow_path(working_folder: str) -> str:
    return os.path.join(working_folder, "files_flow.example.xlsx")


def create_blank_files_flow(path: str) -> None:
    """Create a minimal workbook with Parametrización headers + empty PDF_PDF."""
    param = pd.DataFrame(columns=PARAM_HEADERS)
    # PDF_PDF: row0 = hybrid output name; rows below = "folder, file.pdf"
    pdf_pdf = pd.DataFrame(
        [
            [
                "manifiesto_actuacion.pdf",
                "opinion_sat.pdf",
                "opinion_imss.pdf",
                "opinion_infonavit.pdf",
            ],
            [
                "output, 06_Relaltivo_al_protocolo_de.pdf",
                "output, 11_Cump_de_fiscal_emitida.pdf",
                "output, 14_Cump_de_fiscal_en.pdf",
                "output, 15_Constancia_de_situacion_fiscal.pdf",
            ],
            [
                "fijos, manifiesto.pdf",
                "fijos, opinion_sat.pdf",
                "fijos, opinion_imss.pdf",
                "fijos, opinion_infonavit.pdf",
            ],
        ]
    )
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        param.to_excel(writer, sheet_name=PARAM_SHEET, index=False)
        pdf_pdf.to_excel(writer, sheet_name=PDF_PDF_SHEET, header=False, index=False)


def ensure_files_flow(working_folder: str) -> str:
    """
    Return path to files_flow.xlsx.
    If missing: copy files_flow.example.xlsx, else create a blank template.
    """
    path = files_flow_path(working_folder)
    if os.path.isfile(path):
        return path

    example = example_files_flow_path(working_folder)
    if os.path.isfile(example):
        shutil.copy2(example, path)
        print(f"Created {path} from files_flow.example.xlsx — fill it before running.")
        return path

    create_blank_files_flow(path)
    print(f"Created blank {path} — fill Parametrización and PDF_PDF before running.")
    return path
