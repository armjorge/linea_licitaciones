import os
import shutil
import pandas as pd
from dotenv import load_dotenv

from files_flow_util import PARAM_SHEET, ensure_files_flow

load_dotenv()

ALLOWED_EXTENSIONS = (".pdf", ".xlsx")
FILES_FLOW_SHEET = PARAM_SHEET


def is_allowed_end_file(file_name):
    """Final proposal files must be .pdf or .xlsx."""
    if file_name is None or (isinstance(file_name, float) and pd.isna(file_name)):
        return False
    name = str(file_name).strip()
    if not name or name.lower() == "nan":
        return False
    # Support legacy "subfolder, file.pdf"
    if "," in name:
        name = name.split(",")[-1].strip()
    return name.lower().endswith(ALLOWED_EXTENSIONS)


def normalize_dest_file_name(file_name):
    """Return (optional_subfolder, clean_file_name) from File_name."""
    raw = str(file_name).strip()
    if "," in raw:
        sub_folder, clean = [x.strip() for x in raw.split(",", 1)]
        return sub_folder, clean
    return None, raw


def source_name_candidates(row):
    """Ordered candidate filenames for a row (Source name → letter_name → File_name)."""
    candidates = []
    for col in ("Source name", "letter_name", "File_name"):
        value = row.get(col)
        if pd.isna(value):
            continue
        name = str(value).strip()
        if not name or name.lower() == "nan":
            continue
        if col == "File_name" and "," in name:
            name = name.split(",")[-1].strip()
        if name not in candidates:
            candidates.append(name)
    return candidates


def resolve_source_path(file_manager, row):
    """
    Source file lives at: File management / Source / <source file>.
    Prefer Source name, then letter_name, then File_name (first that exists).
    Returns (found_path_or_None, expected_paths, primary_expected).
    """
    source = row.get("Source")
    if pd.isna(source):
        return None, [], None

    candidates = source_name_candidates(row)
    expected = [
        os.path.join(file_manager, str(source), name) for name in candidates
    ]
    for path in expected:
        if os.path.isfile(path):
            return path, expected, expected[0] if expected else None
    return None, expected, expected[0] if expected else None


def build_dest_path(final_proposal, move_folder, file_name):
    sub_folder, clean_name = normalize_dest_file_name(file_name)
    parts = [final_proposal, str(move_folder)]
    if sub_folder:
        parts.append(sub_folder)
    parts.append(clean_name)
    return os.path.join(*parts)


def evaluate_rows(df, file_manager, final_proposal):
    """
    Classify each row as ready (source exists + valid end File_name) or not.
    Returns list of dicts with status details, including missing_files.
    """
    results = []
    for index, row in df.iterrows():
        requerimiento = row.get("Requerimiento")
        move = row.get("Move")
        file_name = row.get("File_name")

        record = {
            "index": index,
            "id": row.get("id"),
            "Requerimiento": requerimiento if pd.notna(requerimiento) else "",
            "Move": move if pd.notna(move) else "",
            "File_name": file_name if pd.notna(file_name) else "",
            "Source": row.get("Source") if pd.notna(row.get("Source")) else "",
            "ready": False,
            "moved": False,
            "reason": "",
            "source_path": None,
            "dest_path": None,
            "missing_files": [],
        }

        if pd.isna(move) or not str(move).strip():
            record["reason"] = "Missing Move folder in Excel"
            record["missing_files"] = ["Move folder (LEGAL / TECNICO / ECONÓMICO)"]
            results.append(record)
            continue

        if not is_allowed_end_file(file_name):
            record["reason"] = "File_name missing or not .pdf/.xlsx"
            record["missing_files"] = [
                "File_name in Excel ending with .pdf or .xlsx"
            ]
            results.append(record)
            continue

        source_path, expected, primary = resolve_source_path(file_manager, row)
        record["source_path"] = source_path
        record["dest_path"] = build_dest_path(final_proposal, move, file_name)

        if source_path and os.path.isfile(source_path):
            record["ready"] = True
            record["reason"] = "Ready"
        else:
            record["reason"] = "Source file(s) missing"
            # Show the primary expected source; if several aliases were defined, list those missing.
            if primary:
                # Prefer Source name as the "needed" file; include other defined aliases only if distinct.
                names = source_name_candidates(row)
                source_folder = str(row.get("Source"))
                # Primary (first defined) is enough for the user; keep it singular unless letter_name differs.
                missing = [os.path.join(file_manager, source_folder, names[0])]
                record["missing_files"] = missing
            else:
                source_label = str(row.get("Source")) if pd.notna(row.get("Source")) else "?"
                record["missing_files"] = [
                    f"Source name / letter_name under File management/{source_label}/"
                ]
        results.append(record)
    return results


def _short_path(path, file_manager=None):
    """Prefer a short relative path for display."""
    if not path:
        return path
    path = str(path)
    if file_manager and path.startswith(file_manager + os.sep):
        return os.path.join(
            "File management", os.path.relpath(path, file_manager)
        )
    # Keep last 3 parts when absolute
    parts = path.split(os.sep)
    if len(parts) > 3:
        return os.path.join(*parts[-3:])
    return path


def print_success(results, after_move=False, file_manager=None):
    if after_move:
        count = sum(1 for r in results if r["moved"])
        label = "generated"
    else:
        count = sum(1 for r in results if r["ready"])
        label = "ready"

    print("\n" + "=" * 60)
    print(f"SUCCESS — {count} {label} file(s)")
    print("=" * 60)


def print_failure(results, file_manager=None):
    failed = [r for r in results if not r["ready"]]

    print("\n" + "=" * 60)
    print("FAILURE — incomplete / missing")
    print("=" * 60)
    if not failed:
        print("(none)")
        return

    for r in failed:
        target = r["File_name"] or "(no File_name yet)"
        move = r["Move"] or "(no Move)"
        req = str(r["Requerimiento"])
        if len(req) > 80:
            req = req[:77] + "..."

        print(f"\n  File: {target}")
        print(f"  Move: {move}")
        if req:
            print(f"  Requerimiento: {req}")
        print("  Missing to complete:")
        missing = r.get("missing_files") or [r.get("reason") or "Unknown"]
        for item in missing:
            print(f"    - {_short_path(item, file_manager)}")

    print(f"\nTotal failure: {len(failed)}")


def _print_summary_table(df, group_cols, title, only_incomplete=False):
    summary = (
        df.groupby(group_cols, dropna=False)
        .agg(total=("done", "size"), done=("done", "sum"))
        .reset_index()
    )
    summary["not_done"] = summary["total"] - summary["done"]
    summary["pct"] = (summary["done"] / summary["total"] * 100).round(1)

    # Keep overall totals before filtering to incomplete groups.
    all_total = int(summary["total"].sum())
    all_done = int(summary["done"].sum())

    if only_incomplete:
        summary = summary[summary["not_done"] > 0].reset_index(drop=True)

    print("\n" + "=" * 60)
    print(title)
    print("=" * 60)

    if summary.empty:
        print("✅ No incomplete groups.")
        return

    headers = [str(c) for c in group_cols] + ["Done", "Missing", "Total", "%"]
    widths = []
    for col in group_cols:
        widths.append(
            max(12, len(str(col)), int(summary[col].astype(str).str.len().max()))
        )
    widths.extend([6, 8, 7, 7])

    header_line = " ".join(
        f"{h:<{w}}" if i < len(group_cols) else f"{h:>{w}}"
        for i, (h, w) in enumerate(zip(headers, widths))
    )
    print(header_line)
    print("-" * len(header_line))

    for _, row in summary.iterrows():
        parts = []
        for i, col in enumerate(group_cols):
            val = row[col]
            if pd.isna(val) or val == "":
                val = "(empty)"
            parts.append(f"{str(val):<{widths[i]}}")
        parts.append(f"{int(row['done']):>{widths[len(group_cols)]}}")
        parts.append(f"{int(row['not_done']):>{widths[len(group_cols)+1]}}")
        parts.append(f"{int(row['total']):>{widths[len(group_cols)+2]}}")
        parts.append(f"{row['pct']:>{widths[len(group_cols)+3]-1}.1f}%")
        print(" ".join(parts))

    pct = (all_done / all_total * 100) if all_total else 0
    print("-" * len(header_line))
    lead = f"{'ALL':<{widths[0]}}"
    for w in widths[1:len(group_cols)]:
        lead += " " + (" " * w)
    print(
        f"{lead} "
        f"{all_done:>{widths[len(group_cols)]}} "
        f"{all_total - all_done:>{widths[len(group_cols)+1]}} "
        f"{all_total:>{widths[len(group_cols)+2]}} "
        f"{pct:>{widths[len(group_cols)+3]-1}.1f}%"
    )


def print_completeness(results, after_move=False):
    scoped = [r for r in results if r["Move"]]
    if not scoped:
        print("\nNo rows with Move to summarize.")
        return

    df = pd.DataFrame(scoped)
    df["done"] = df["moved"] if after_move else df["ready"]
    # Normalize empty Source for grouping display
    df["Source"] = df["Source"].fillna("").replace("", "(empty)")

    _print_summary_table(
        df, ["Move"], "COMPLETENESS BY Move (incomplete only)", only_incomplete=True
    )
    _print_summary_table(
        df,
        ["Move", "Source"],
        "COMPLETENESS BY Move × Source (incomplete only)",
        only_incomplete=True,
    )


def print_report(results, file_manager, after_move=False):
    print_success(results, after_move=after_move, file_manager=file_manager)
    print_failure(results, file_manager=file_manager)
    print_completeness(results, after_move=after_move)


def clear_final_proposal(final_proposal, move_folders=None):
    """
    Wipe final_proposal completely and recreate the root + Move subfolders.
    Guarantees no leftover files from a previous run.
    """
    if os.path.exists(final_proposal):
        before_files = []
        for root, _dirs, files in os.walk(final_proposal):
            for name in files:
                before_files.append(os.path.join(root, name))
        shutil.rmtree(final_proposal)
        print(
            f"🧹 Removed final proposal folder "
            f"({len(before_files)} old file(s) deleted)"
        )
    else:
        print("🧹 Final proposal folder did not exist; creating fresh.")

    os.makedirs(final_proposal, exist_ok=True)

    created = []
    for folder in sorted({str(f).strip() for f in (move_folders or []) if pd.notna(f) and str(f).strip()}):
        path = os.path.join(final_proposal, folder)
        os.makedirs(path, exist_ok=True)
        created.append(folder)

    if created:
        print(f"📁 Recreated Move folders: {', '.join(created)}")
    print(f"✅ Clean final proposal ready at: {final_proposal}")


def move_ready_files(results):
    """Copy ready files into final_proposal using File_name as destination name."""
    ready = [r for r in results if r["ready"]]
    if not ready:
        return results

    print(f"\n🚀 Copying {len(ready)} ready file(s)...")
    for r in ready:
        dest = r["dest_path"]
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        try:
            shutil.copy2(r["source_path"], dest)
            r["moved"] = True
        except Exception as e:
            r["moved"] = False
            r["ready"] = False
            r["reason"] = f"Copy failed: {e}"
            r["missing_files"] = [str(e)]
    return results


def verify_final_proposal(final_proposal, results):
    """
    Confirm final_proposal only contains files from this move run.
    """
    moved_paths = {
        os.path.normpath(r["dest_path"])
        for r in results
        if r.get("moved") and r.get("dest_path")
    }

    present = []
    for root, _dirs, files in os.walk(final_proposal):
        for name in files:
            present.append(os.path.normpath(os.path.join(root, name)))

    present_set = set(present)
    unexpected = sorted(present_set - moved_paths)
    missing_moved = sorted(moved_paths - present_set)

    print("\n" + "=" * 60)
    print("FINAL FOLDER VERIFICATION")
    print("=" * 60)
    print(f"Moved this run : {len(moved_paths)}")
    print(f"Files on disk  : {len(present_set)}")

    if unexpected:
        print(f"\n⚠️ Unexpected leftover file(s) ({len(unexpected)}):")
        for path in unexpected:
            print(f"   - {os.path.relpath(path, final_proposal)}")
    if missing_moved:
        print(f"\n❌ Reported as moved but missing on disk ({len(missing_moved)}):")
        for path in missing_moved:
            print(f"   - {os.path.relpath(path, final_proposal)}")
    if not unexpected and not missing_moved:
        print("✅ Final folder contains only files from this move.")


def load_context():
    working_folder = os.getenv("working_folder")
    final_proposal = os.getenv("final_proposal")
    if not working_folder or not final_proposal:
        raise RuntimeError("Set working_folder and final_proposal in .env")

    file_manager = os.path.join(working_folder, "File management")
    excel_file = ensure_files_flow(working_folder)
    df = pd.read_excel(excel_file, sheet_name=FILES_FLOW_SHEET)
    return working_folder, file_manager, final_proposal, excel_file, df


def run_check(file_manager, final_proposal, df):
    results = evaluate_rows(df, file_manager, final_proposal)
    print_report(results, file_manager, after_move=False)
    return results


def run_move(file_manager, final_proposal, df):
    results = evaluate_rows(df, file_manager, final_proposal)
    ready_count = sum(1 for r in results if r["ready"])

    print_report(results, file_manager, after_move=False)

    if ready_count == 0:
        return results

    print(
        "\n⚠️ This will DELETE everything inside final_proposal, "
        "recreate Move folders, and copy only the ready files from this run."
    )
    answer = input(
        f"Wipe '{final_proposal}' and move {ready_count} ready file(s)? [y/N]: "
    ).strip().lower()
    if answer not in ("y", "yes", "s", "si", "sí"):
        print("Cancelled.")
        return results

    move_folders = df["Move"].dropna().unique().tolist() if "Move" in df.columns else []
    clear_final_proposal(final_proposal, move_folders=move_folders)
    results = move_ready_files(results)
    verify_final_proposal(final_proposal, results)
    print("\n>>> After move:")
    print_report(results, file_manager, after_move=True)
    print(f"\n✨ Final proposal folder: {final_proposal}")
    return results


def print_menu(working_folder, final_proposal):
    print("\n" + "=" * 60)
    print("MUEVE PROPUESTA")
    print("=" * 60)
    print(f"working_folder : {working_folder}")
    print(f"final_proposal : {final_proposal}")
    print("-" * 60)
    print("1. Check readiness & completeness summary")
    print("2. Wipe final_proposal and move ready files")
    print("0. Exit")
    print("=" * 60)


def main():
    try:
        working_folder, file_manager, final_proposal, excel_file, df = load_context()
    except Exception as e:
        print(f"❌ Setup error: {e}")
        return

    print(f"Loaded {len(df)} rows from {excel_file}")

    while True:
        print_menu(working_folder, final_proposal)
        choice = input("Choose an option: ").strip()

        if choice == "1":
            run_check(file_manager, final_proposal, df)
        elif choice == "2":
            run_move(file_manager, final_proposal, df)
        elif choice == "0":
            print("Bye.")
            break
        else:
            print("Invalid choice. Please select 0-2.")
            continue

        input("\nPress Enter to return to menu...")


if __name__ == "__main__":
    main()
