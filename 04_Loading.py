import os
import shutil
import subprocess
import sys

import pandas as pd
from dotenv import load_dotenv

from files_flow_util import PARAM_SHEET, ensure_files_flow

load_dotenv()

FILES_FLOW_SHEET = PARAM_SHEET


def load_parametrizacion():
    working_folder = os.getenv("working_folder")
    if not working_folder:
        raise RuntimeError("Set working_folder in .env")

    excel_file = ensure_files_flow(working_folder)
    df = pd.read_excel(excel_file, sheet_name=FILES_FLOW_SHEET)
    if "Move" not in df.columns or "File_name" not in df.columns or "id" not in df.columns:
        raise RuntimeError(
            "Sheet Parametrización must include columns: id, File_name, Move"
        )
    return working_folder, excel_file, df


def copy_to_clipboard(text):
    """Copy text to the system clipboard (Linux/macOS/Windows/OrbStack)."""
    text = str(text)
    if sys.platform == "win32":
        # Windows clip.exe expects UTF-16LE without BOM for best compatibility.
        completed = subprocess.run(
            ["clip"],
            input=text.encode("utf-16le"),
            check=False,
        )
        if completed.returncode == 0:
            return True
        raise RuntimeError("Windows clip.exe failed")

    candidates = [
        ["pbcopy"],  # macOS / OrbStack bridge
        ["wl-copy"],
        ["xclip", "-selection", "clipboard"],
        ["xsel", "--clipboard", "--input"],
    ]
    for cmd in candidates:
        if not shutil.which(cmd[0]):
            continue
        completed = subprocess.run(
            cmd,
            input=text.encode("utf-8"),
            check=False,
        )
        if completed.returncode == 0:
            return True

    raise RuntimeError(
        "No clipboard tool found. Install one of: pbcopy, wl-copy, xclip, xsel"
    )


def check_ids(df_filtered):
    """
    Ensure id runs 1..N with no gaps/duplicates after sorting.
    Returns (ok, message).
    """
    ids = pd.to_numeric(df_filtered["id"], errors="coerce")
    if ids.isna().any():
        bad = df_filtered.loc[ids.isna(), ["id", "File_name"]]
        return False, f"Non-numeric or empty id values found:\n{bad.to_string(index=False)}"

    ids_list = ids.astype(int).tolist()
    expected = list(range(1, len(ids_list) + 1))
    if ids_list == expected:
        return True, f"id sequence OK: 1 → {len(ids_list)}"

    return (
        False,
        f"id sequence invalid.\n"
        f"  Expected: {expected}\n"
        f"  Found:    {ids_list}",
    )


def choose_move(df):
    moves = (
        df["Move"]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .unique()
        .tolist()
    )
    if not moves:
        print("❌ No Move values found in Excel.")
        return None

    print("\nAvailable Move options:")
    for i, move in enumerate(moves, start=1):
        count = len(df[df["Move"].astype(str).str.strip() == move])
        print(f"  {i}. {move}  ({count} file(s))")

    raw = input("Choose Move number (or 0 to cancel): ").strip()
    if raw == "0":
        return None
    try:
        idx = int(raw)
    except ValueError:
        print("Invalid choice.")
        return None
    if idx < 1 or idx > len(moves):
        print("Invalid choice.")
        return None
    return moves[idx - 1]


def run_loading_session(df):
    chosen = choose_move(df)
    if not chosen:
        print("Cancelled.")
        return

    df_filtered = df[df["Move"].astype(str).str.strip() == chosen].copy()
    df_filtered = df_filtered.sort_values("id", kind="mergesort").reset_index(drop=True)

    print(f"\nSelected Move: {chosen}")
    print(f"Rows: {len(df_filtered)}")

    ok, message = check_ids(df_filtered)
    print(message)
    if not ok:
        answer = input("Continue anyway? [y/N]: ").strip().lower()
        if answer not in ("y", "yes", "s", "si", "sí"):
            print("Cancelled.")
            return

    total = len(df_filtered)
    for i, row in df_filtered.iterrows():
        file_name = row.get("File_name")
        row_id = row.get("id")

        if pd.isna(file_name) or not str(file_name).strip():
            print(f"\n[{i + 1}/{total}] id={row_id} — skipped (empty File_name)")
            input("Press Enter for next...")
            continue

        file_name = str(file_name).strip()
        try:
            copy_to_clipboard(file_name)
            clip_note = "copied to clipboard"
        except Exception as e:
            clip_note = f"clipboard failed ({e})"

        print("\n" + "-" * 60)
        print(f"[{i + 1}/{total}] Move={chosen}  id={row_id}")
        print(f"File_name: {file_name}")
        print(f"Clipboard: {clip_note}")
        print("-" * 60)
        input("Press Enter for next file...")

    print(f"\n✨ Done with Move '{chosen}'. Returning to menu.")


def print_menu(working_folder, excel_file):
    print("\n" + "=" * 60)
    print("LOADING — clipboard File_name helper")
    print("=" * 60)
    print(f"working_folder : {working_folder}")
    print(f"files_flow     : {excel_file}")
    print("-" * 60)
    print("1. Load by Move (copy File_name one by one)")
    print("0. Exit")
    print("=" * 60)


def main():
    try:
        working_folder, excel_file, df = load_parametrizacion()
    except Exception as e:
        print(f"❌ Setup error: {e}")
        return

    print(f"Loaded {len(df)} rows from {excel_file}")

    while True:
        print_menu(working_folder, excel_file)
        choice = input("Choose an option: ").strip()

        if choice == "1":
            # Reload Excel each session so edits are picked up.
            try:
                _, excel_file, df = load_parametrizacion()
            except Exception as e:
                print(f"❌ Error reloading Excel: {e}")
                continue
            run_loading_session(df)
        elif choice == "0":
            print("Bye.")
            break
        else:
            print("Invalid choice. Please select 0-1.")
            continue


if __name__ == "__main__":
    main()
