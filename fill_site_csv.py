"""
fill_site_csv.py

Fills the online submission site's per-class CSV templates using the NAMA
and NA columns from the "Pengolahan Nilai Rapor" workbook produced by
na_report.py.

Matching logic (same as the earlier grade-filler project):
  - Site template columns: NISN, Nama, Nilai (NISN already filled in, Nilai blank)
  - Match students by name (case-insensitive, whitespace-trimmed)
  - A student on the site template but missing from the report ("absent",
    or enrolled in a different class online) gets a blank Nilai, and is
    listed in the mismatch report instead of guessed at
  - A student in the report but not on that class's site template is
    also listed in the mismatch report (they may belong to a different
    class online)
  - Output file name: same as the template but with "Template_" removed,
    e.g. "Template_Nilai_PTS_X TJKT 1.csv" -> "Nilai_PTS_X TJKT 1.csv"

Usage:
    python fill_site_csv.py report.xlsx templates_dir/ output_dir/

report.xlsx    -- output of na_report.py (one sheet per class, NAMA in
                   column B, NA in the "NA" header column)
templates_dir/ -- folder containing the site's Template_Nilai_*.csv files,
                  one per class. The class name is matched against the
                  part of the filename after the last underscore-prefixed
                  segment (e.g. "X TJKT 1" in "Template_Nilai_PTS_X TJKT 1.csv")
                  falling back to substring matching against report sheet names.
output_dir/    -- where the filled CSVs (and mismatch report) are written
"""

import sys
import csv
import re
from pathlib import Path
from openpyxl import load_workbook


def norm(name):
    return re.sub(r"\s+", " ", str(name).strip()).upper()


def read_report(report_path):
    """Return {class_name: {normalized_name: na_value}}"""
    wb = load_workbook(report_path, data_only=True)
    classes = {}
    for sheet_name in wb.sheetnames:
        if sheet_name.strip().lower() == "peringatan":
            continue
        ws = wb[sheet_name]
        header_row = None
        for r in range(1, ws.max_row + 1):
            vals = [c.value for c in ws[r]]
            if "NAMA" in vals and "NA" in vals:
                header_row = r
                break
        if header_row is None:
            continue
        headers = [c.value for c in ws[header_row]]
        no_col = headers.index("NO") + 1 if "NO" in headers else None
        nama_col = headers.index("NAMA") + 1
        na_col = headers.index("NA") + 1

        students = {}
        for r in range(header_row + 1, ws.max_row + 1):
            nama = ws.cell(row=r, column=nama_col).value
            na = ws.cell(row=r, column=na_col).value
            if nama is None or str(nama).strip() == "":
                continue
            # footer rows (signature block, formula note) have no numeric NO --
            # stop reading students as soon as we hit one
            if no_col is not None:
                no_val = ws.cell(row=r, column=no_col).value
                if not isinstance(no_val, (int, float)):
                    break
            students[norm(nama)] = na
        classes[sheet_name] = students
    return classes


def match_class_name(template_filename, class_names):
    """Find which report class this template file belongs to."""
    stem = Path(template_filename).stem
    # exact substring match first (longest class name wins, avoids "X TJKT 1" matching "X TJKT 10")
    candidates = [c for c in class_names if c.upper() in stem.upper()]
    if candidates:
        return max(candidates, key=len)
    return None


def fill_templates(report_path, templates_dir, output_dir):
    classes = read_report(report_path)
    class_names = list(classes.keys())
    templates_dir = Path(templates_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    mismatches = []
    template_files = sorted(templates_dir.glob("*.csv"))
    if not template_files:
        print(f"Tidak ada file .csv di {templates_dir}")
        return

    for tpl_path in template_files:
        class_name = match_class_name(tpl_path.name, class_names)
        if class_name is None:
            mismatches.append(f"{tpl_path.name}: tidak cocok dengan kelas manapun di laporan, dilewati")
            continue

        report_students = dict(classes[class_name])  # copy so we can mark used
        matched_names = set()

        with open(tpl_path, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            fieldnames = reader.fieldnames
            rows = list(reader)

        if not fieldnames or "Nama" not in fieldnames or "Nilai" not in fieldnames:
            mismatches.append(f"{tpl_path.name}: kolom 'Nama'/'Nilai' tidak ditemukan, dilewati")
            continue

        for row in rows:
            key = norm(row["Nama"])
            if key in report_students:
                row["Nilai"] = report_students[key]
                matched_names.add(key)
            else:
                row["Nilai"] = ""
                mismatches.append(
                    f"{class_name} / {tpl_path.name}: '{row['Nama']}' ada di template situs tapi tidak ada di laporan (tidak hadir / beda kelas?)"
                )

        # students in the report but not on the site template for this class
        for key, na in report_students.items():
            if key not in matched_names:
                mismatches.append(
                    f"{class_name} / {tpl_path.name}: '{key.title()}' ada di laporan tapi tidak ditemukan di template situs untuk kelas ini"
                )

        out_name = tpl_path.name
        if out_name.startswith("Template_"):
            out_name = out_name[len("Template_"):]
        out_path = output_dir / out_name

        with open(out_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)

        print(f"{tpl_path.name} -> {out_path.name} ({len(matched_names)}/{len(rows)} siswa cocok)")

    if mismatches:
        mismatch_path = output_dir / "Peringatan_Pencocokan.txt"
        with open(mismatch_path, "w", encoding="utf-8") as f:
            f.write("\n".join(mismatches))
        print(f"\n{len(mismatches)} peringatan pencocokan. Lihat {mismatch_path.name}")
    else:
        print("\nTidak ada peringatan pencocokan.")


def main():
    if len(sys.argv) != 4:
        print("Usage: python fill_site_csv.py report.xlsx templates_dir/ output_dir/")
        sys.exit(1)
    fill_templates(sys.argv[1], sys.argv[2], sys.argv[3])


if __name__ == "__main__":
    main()
