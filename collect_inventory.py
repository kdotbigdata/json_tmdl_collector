#!/usr/bin/env python3
"""
Inventory collector for PBIP projects (strict folder names).

Usage:
  python collect_inventory.py --root /path/to/root

Expects each project folder under --root to contain one PBIP export, either
directly or inside a dedicated subfolder (e.g., ``dashboard/``) with:
  - A pbip file (e.g., dashboard.pbip)
  - A folder: dashboard.Report
  - A folder: dashboard.SemanticModel

Copies JSON/TMDL files into an inventory directory located ONE LEVEL UP from root:

  ../inventory/
    manifest.csv
    report/
      json_files/
      tmdl_files/
    semanticmodel/
      json_files/
      tmdl_files/

Output filenames are <pbipbase>_<originalname>.<ext>, e.g., dashboard_pages.json.
"""

import argparse
import csv
import shutil
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

def find_first_pbip(project_dir: Path) -> Optional[Path]:
    """Return the first PBIP file in ``project_dir`` or its descendants.

    Prefers PBIP files located directly in ``project_dir``. Falls back to the
    first PBIP in subdirectories where the standard Report/SemanticModel
    folders live alongside the PBIP file.
    """

    direct = sorted(project_dir.glob("*.pbip"))
    if direct:
        return direct[0]

    nested = sorted(
        p for p in project_dir.rglob("*.pbip")
        if (p.parent / f"{p.stem}.Report").is_dir()
        and (p.parent / f"{p.stem}.SemanticModel").is_dir()
    )
    if nested:
        return nested[0]

    # Nothing with sibling folders; fall back to any PBIP found.
    fallback = sorted(project_dir.rglob("*.pbip"))
    return fallback[0] if fallback else None

WINDOWS_FORBIDDEN = set('<>:"/\\|?*')  # Reserved characters on Windows paths

def sanitize_pbip_base(stem: str) -> str:
    """Return a file-name-safe base derived from the PBIP stem."""
    cleaned_chars = []
    for ch in stem.strip():
        if ch in WINDOWS_FORBIDDEN or ord(ch) < 32:
            cleaned_chars.append('_')
        elif ch.isspace():
            cleaned_chars.append('_')
        else:
            cleaned_chars.append(ch)
    cleaned = ''.join(cleaned_chars).rstrip('._-')
    return cleaned or 'pbip'

def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)

def next_available_path(target: Path) -> Path:
    """If target exists, add -1, -2, ... before suffix."""
    if not target.exists():
        return target
    stem, suffix = target.stem, target.suffix
    parent = target.parent
    i = 1
    while True:
        candidate = parent / f"{stem}-{i}{suffix}"
        if not candidate.exists():
            return candidate
        i += 1

def gather_candidates(pbip_root: Path, pbip: Path) -> Dict[str, List[Path]]:
    """Return candidate files from folders matching the PBIP stem.

    Power BI Desktop exports names folders after the PBIP file's stem, e.g.
    ``Report.pbip`` -> ``Report.Report`` / ``Report.SemanticModel``. Some
    exports may include spaces or punctuation, so we derive the expected folder
    names from the actual PBIP filename and fall back to any ``*.Report`` /
    ``*.SemanticModel`` directories if the direct match is missing.
    """
    out = {k: [] for k in ("report_json", "report_tmdl", "sem_json", "sem_tmdl")}

    expected_report = pbip_root / f"{pbip.stem}.Report"
    expected_sem    = pbip_root / f"{pbip.stem}.SemanticModel"

    report_roots: List[Path] = []
    sem_roots: List[Path] = []

    if expected_report.is_dir():
        report_roots.append(expected_report)
    else:
        report_roots.extend(p for p in pbip_root.glob("*.Report") if p.is_dir())

    if expected_sem.is_dir():
        sem_roots.append(expected_sem)
    else:
        sem_roots.extend(p for p in pbip_root.glob("*.SemanticModel") if p.is_dir())

    for report_root in report_roots:
        out["report_json"].extend(report_root.rglob("*.json"))
        out["report_tmdl"].extend(report_root.rglob("*.tmdl"))

    for sem_root in sem_roots:
        out["sem_json"].extend(sem_root.rglob("*.json"))
        out["sem_tmdl"].extend(sem_root.rglob("*.tmdl"))

    return out

def build_dest_name(pbip_base: str, src: Path) -> str:
    """
    <pbipbase>_<originalfilename> with extension preserved.
    Example: dashboard + pages.json -> dashboard_pages.json
    """
    return f"{pbip_base}_{src.stem}{src.suffix}"

def collect(
    root: Path,
    dry_run: bool = False,
    verbose: bool = True,
) -> None:
    # Inventory lives one level up from root
    inv_base = root.parent / "inventory"

    # Output folders
    report_json_dir = inv_base / "report" / "json_files"
    report_tmdl_dir = inv_base / "report" / "tmdl_files"
    sem_json_dir    = inv_base / "semanticmodel" / "json_files"
    sem_tmdl_dir    = inv_base / "semanticmodel" / "tmdl_files"
    manifest_path   = inv_base / "manifest.csv"

    # Make dirs
    for d in (report_json_dir, report_tmdl_dir, sem_json_dir, sem_tmdl_dir):
        ensure_dir(d)

    # Iterate immediate subdirectories as projects
    projects = [p for p in root.iterdir() if p.is_dir()]
    if verbose:
        print(f"Found {len(projects)} project folder(s) under {root}")

    # Prepare manifest rows
    manifest_rows: List[Tuple[str, str]] = []  # (project_rel, pbip_name)

    for project in projects:
        pbip = find_first_pbip(project)
        if not pbip:
            if verbose:
                print(f"Skipping {project.name}: no pbip found.")
            continue

        pbip_base = sanitize_pbip_base(pbip.stem)
        try:
            project_rel = str(pbip.parent.relative_to(root))
        except ValueError:
            # Fallback if PBIP resides outside root (shouldn't happen in normal usage)
            project_rel = str(project.relative_to(root))
        manifest_rows.append((project_rel, pbip.name))

        # Locate candidate files using folder names derived from the PBIP file
        cand = gather_candidates(pbip.parent, pbip)

        def copy_many(files: Iterable[Path], dest_dir: Path):
            for src in files:
                if not src.exists():
                    if verbose:
                        print(f"Skip (missing): {src}")
                    continue
                if not src.is_file():
                    if verbose:
                        print(f"Skip (not a file): {src}")
                    continue
                dst_name = build_dest_name(pbip_base, src)
                dst_path = dest_dir / dst_name
                dst_path = next_available_path(dst_path)
                if verbose:
                    print(f"Copy: {src} -> {dst_path}")
                if not dry_run:
                    ensure_dir(dst_path.parent)
                    shutil.copy2(src, dst_path)

        copy_many(cand["report_json"], report_json_dir)
        copy_many(cand["report_tmdl"], report_tmdl_dir)
        copy_many(cand["sem_json"],    sem_json_dir)
        copy_many(cand["sem_tmdl"],    sem_tmdl_dir)

    # Write manifest
    if verbose:
        print(f"Writing manifest: {manifest_path}")
    if not dry_run:
        ensure_dir(inv_base)
        with manifest_path.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["project_folder", "pbip_file"])
            w.writerows(manifest_rows)

def main():
    ap = argparse.ArgumentParser(description="Collect JSON/TMDL inventory from PBIP projects (strict).")
    ap.add_argument("--root", type=Path, default=Path.cwd(),
                    help="Root directory containing project folders (default: CWD).")
    ap.add_argument("--dry-run", action="store_true", help="Show actions without copying.")
    ap.add_argument("--quiet", action="store_true", help="Reduce output.")
    args = ap.parse_args()

    if not args.root.exists() or not args.root.is_dir():
        raise SystemExit(f"Root path not found or not a directory: {args.root}")

    collect(args.root, dry_run=args.dry_run, verbose=not args.quiet)

if __name__ == "__main__":
    main()
