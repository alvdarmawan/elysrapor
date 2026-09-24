"""
na_report.py

Reads per-class grade sheets (Name, Tugas dan Latihan, Nilai 1-8, Ulangan
Harian, PTS, PAS) and produces a "Pengolahan Nilai Rapor" workbook in the
school's official layout, with Rata-rata NH and NA computed per student.

Usage:
    python na_report.py input.xlsx output.xlsx
"""

import sys
import copy
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# CONFIG — edit these to change behaviour without touching the logic below
# ----------------------------------------------------------------------

# Letter grade -> extra points
LETTER_POINTS = {"A": 3, "B+": 2, "B": 1}

# Weights when the class has NO Ulangan Harian column filled in
WEIGHTS_NO_UH = {"harian": 0.60, "pts": 0.40, "pas": 0} #Activate this if mid-sem
#WEIGHTS_NO_UH = {"harian": 0.40, "pts": 0.30, "pas": 0.30} #Activate ths if end-sem

# Weights when the class HAS Ulangan Harian
WEIGHTS_UH = {"harian": 0.30, "uh": 0.45, "pts": 0.25, "pas": 0} #Activate this if mid-sem
#WEIGHTS_UH = {"harian": 0.20, "uh": 0.35, "pts": 0.20, "pas": 0.25} #Activate this if end-sem

# Minimum number of CP (harian) columns to show, even if the class uses fewer
MIN_CP_COLUMNS = 6

# Header/footer text
SCHOOL_NAME = "SMK BINA TARUNA"
ACADEMIC_YEAR = "TAHUN AJARAN 2025/2026"
MATA_PELAJARAN = "BAHASA INGGRIS"
SEMESTER = ""
TEACHER_NAME = "Elys Putri Karismawati"

INPUT_HEADERS = {
    "nama": "Nama",
    "tugas": "Tugas dan Latihan",
    "harian": [f"Nilai {i}" for i in range(1, 9)],
    "uh": "Ulangan Harian",
    "pts": "PTS",
    "pas": "PAS",
}

# ----------------------------------------------------------------------
# Step 1 — read the input workbook
# ----------------------------------------------------------------------

def parse_harian_cell(value, warnings, ctx):
    """Return (numeric_value_or_None, extra_points) for one Nilai cell."""
    if value is None or (isinstance(value, str) and value.strip() == ""):
        return None, 0
    if isinstance(value, (int, float)):
        return float(value), 0
    text = str(value).strip().upper()
    # split things like "AA", "AB+", "A B" into individual letter tokens
    tokens = []
    i = 0
    while i < len(text):
        if text[i:i+2] == "B+":
            tokens.append("B+")
            i += 2
        elif text[i] in ("A", "B"):
            tokens.append(text[i])
            i += 1
        elif text[i] in (" ", ","):
            i += 1
        else:
            warnings.append(f"{ctx}: nilai tidak dikenali '{value}', diabaikan")
            return None, 0
    extra = 0
    for t in tokens:
        if t not in LETTER_POINTS:
            warnings.append(f"{ctx}: huruf tidak dikenali '{t}' dalam '{value}', diabaikan")
            continue
        extra += LETTER_POINTS[t]
    return None, extra


def read_input(path):
    wb = load_workbook(path, data_only=True)
    classes = {}
    warnings = []

    for sheet_name in wb.sheetnames:
        if sheet_name.strip().lower() in ("petunjuk", "instructions", "readme"):
            continue
        ws = wb[sheet_name]
        header_row = [c.value for c in ws[1]]
        col_index = {h: idx for idx, h in enumerate(header_row) if h}

        required = [INPUT_HEADERS["nama"], INPUT_HEADERS["pts"], INPUT_HEADERS["pas"]]
        if not all(r in col_index for r in required):
            warnings.append(f"Sheet '{sheet_name}': dilewati, header wajib tidak lengkap")
            continue

        harian_cols = [h for h in INPUT_HEADERS["harian"] if h in col_index]
        tugas_col = col_index.get(INPUT_HEADERS["tugas"])
        uh_col = col_index.get(INPUT_HEADERS["uh"])
        pts_col = col_index[INPUT_HEADERS["pts"]]
        pas_col = col_index[INPUT_HEADERS["pas"]]
        nama_col = col_index[INPUT_HEADERS["nama"]]

        students = []
        class_has_uh = False

        for row in ws.iter_rows(min_row=2, values_only=False):
            nama = row[nama_col].value
            if nama is None or str(nama).strip() == "":
                continue
            nama = str(nama).strip()
            ctx = f"{sheet_name} / {nama}"

            numeric_grades = []
            extra_points = 0
            for h in harian_cols:
                raw = row[col_index[h]].value
                num, extra = parse_harian_cell(raw, warnings, ctx)
                if num is not None:
                    numeric_grades.append(num)
                extra_points += extra

            tugas_val = 0
            if tugas_col is not None:
                raw_t = row[tugas_col].value
                if isinstance(raw_t, (int, float)):
                    tugas_val = float(raw_t)
                elif raw_t not in (None, ""):
                    warnings.append(f"{ctx}: nilai tugas dan latihan '{raw_t}' bukan angka, diabaikan")

            uh_val = None
            if uh_col is not None:
                raw_uh = row[uh_col].value
                if isinstance(raw_uh, (int, float)):
                    uh_val = float(raw_uh)
                    class_has_uh = True

            pts_raw = row[pts_col].value
            pas_raw = row[pas_col].value
            if pts_raw is None:
                warnings.append(f"{ctx}: PTS kosong, dianggap 0")
            if pas_raw is None:
                warnings.append(f"{ctx}: PAS kosong, dianggap 0")
            pts_val = float(pts_raw) if isinstance(pts_raw, (int, float)) else 0.0
            pas_val = float(pas_raw) if isinstance(pas_raw, (int, float)) else 0.0

            if not numeric_grades and extra_points == 0 and tugas_val == 0:
                warnings.append(f"{ctx}: tidak ada nilai harian numerik maupun huruf")

            students.append({
                "nama": nama,
                "numeric_grades": numeric_grades,
                "extra_points": extra_points,
                "tugas": tugas_val,
                "uh": uh_val,
                "pts": pts_val,
                "pas": pas_val,
            })

        # second pass: any student missing UH in a class that has UH -> 0 + warning
        if class_has_uh:
            for s in students:
                if s["uh"] is None:
                    warnings.append(f"{sheet_name} / {s['nama']}: kelas memakai UH tapi UH kosong, dianggap 0")
                    s["uh"] = 0.0

        classes[sheet_name] = {
            "students": students,
            "has_uh": class_has_uh,
            "n_harian_cols": len(harian_cols),
        }

    return classes, warnings


# ----------------------------------------------------------------------
# Step 2 — compute Rata-rata NH and NA
# ----------------------------------------------------------------------

def compute(classes):
    for cname, cdata in classes.items():
        weights = WEIGHTS_UH if cdata["has_uh"] else WEIGHTS_NO_UH
        for s in cdata["students"]:
            numeric = s["numeric_grades"]
            avg_no_uh = sum(numeric) / len(numeric) if numeric else 0.0
            harian_for_na = avg_no_uh + s["extra_points"] + s["tugas"]

            if cdata["has_uh"]:
                all_vals = numeric + [s["uh"]]
                avg_with_uh = sum(all_vals) / len(all_vals) if all_vals else 0.0
                rata_nh = avg_with_uh + s["extra_points"] + s["tugas"]
                na = (harian_for_na * weights["harian"]
                      + s["uh"] * weights["uh"]
                      + s["pts"] * weights["pts"]
                      + s["pas"] * weights["pas"])
            else:
                rata_nh = harian_for_na
                na = (harian_for_na * weights["harian"]
                      + s["pts"] * weights["pts"]
                      + s["pas"] * weights["pas"])

            #s["harian_for_na"] = harian_for_na
            #s["rata_nh"] = round(rata_nh)
            #s["na"] = round(na)
            s["harian_for_na"] = harian_for_na
            s["rata_nh"] = round(min(rata_nh, 100))
            s["na"] = round(min(na, 100))
    return classes


# ----------------------------------------------------------------------
# Step 3 — write the output workbook in the official layout
# ----------------------------------------------------------------------

THIN = Side(style="thin")
BORDER_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FONT_TITLE = Font(name="Calibri", size=14, bold=True)
FONT_HDR = Font(name="Calibri", size=11, bold=False)
FONT_BODY = Font(name="Calibri", size=11)
FONT_NAMA = Font(name="Book Antiqua", size=10)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_LEFT = Alignment(horizontal="left", vertical="top")


def write_class_sheet(wb, sheet_name, cdata):
    ws = wb.create_sheet(sheet_name)
    #n_cp = max(MIN_CP_COLUMNS, cdata["n_harian_cols"])
    n_cp = 6

    # layout columns: A=NO, B=NAMA, C..(C+n_cp-1)=CP, then Rata-rata NH, PSTS, PSAT, NA, KET
    col_cp_start = 3  # C
    col_cp_end = col_cp_start + n_cp - 1
    col_rata = col_cp_end + 1
    col_pts = col_rata + 1
    col_pas = col_pts + 1
    col_na = col_pas + 1
    col_ket = col_na + 1
    last_col = col_ket

    def L(idx):
        return get_column_letter(idx)

    # ---- title block ----
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=last_col)
    ws["A1"] = "PENGOLAHAN NILAI RAPOR"
    ws["A2"] = SCHOOL_NAME
    ws["A2"].font = Font(bold=True)
    ws["A3"] = ACADEMIC_YEAR
    ws["A3"].font = Font(bold=True)
    for r in (1, 2, 3):
        ws[f"A{r}"].font = FONT_TITLE
        ws[f"A{r}"].alignment = ALIGN_CENTER

    ws["A5"] = f"MATA PELAJARAN : {MATA_PELAJARAN}"
    ws["A5"].font = Font(bold=True)
    ws["A6"] = f"KELAS                      : {sheet_name}"
    ws["A6"].font = Font(bold=True)
    ws[f"{L(col_pts)}6"] = f"Semester : {SEMESTER}"

    # ---- table header (rows 7-8) ----
    ws.merge_cells(start_row=7, start_column=1, end_row=8, end_column=1)   # NO
    ws.merge_cells(start_row=7, start_column=2, end_row=8, end_column=2)   # NAMA
    ws.merge_cells(start_row=7, start_column=col_cp_start, end_row=7, end_column=col_cp_end)  # CP
    for col in (col_rata, col_pts, col_pas, col_na, col_ket):
        ws.merge_cells(start_row=7, start_column=col, end_row=8, end_column=col)

    ws.cell(row=7, column=1, value="NO")
    ws.cell(row=7, column=2, value="NAMA")
    ws.cell(row=7, column=col_cp_start, value="CP")
    ws.cell(row=7, column=col_rata, value="Rata-rata NH")
    ws.cell(row=7, column=col_pts, value="PSTS")
    ws.cell(row=7, column=col_pas, value="PSAS")
    ws.cell(row=7, column=col_na, value="NA")
    ws.cell(row=7, column=col_ket, value="KET")
    for i in range(n_cp):
        ws.cell(row=8, column=col_cp_start + i, value=i + 1)

    for r in (7, 8):
        for c in range(1, last_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = FONT_HDR
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_ALL
        ws.row_dimensions[r].height = 14.25

    # ---- student rows ----
    students = cdata["students"]
    start_row = 9
    for i, s in enumerate(students):
        r = start_row + i
        ws.cell(row=r, column=1, value=i + 1)
        #ws.cell(row=r, column=2, value=s["nama"])
        ws.cell(row=r, column=2, value=str(s["nama"]).upper())
        for j, val in enumerate(s["numeric_grades"][:n_cp]):
            ws.cell(row=r, column=col_cp_start + j, value=val)
        ws.cell(row=r, column=col_rata, value=s["rata_nh"])
        ws.cell(row=r, column=col_pts, value=round(s["pts"]))
        ws.cell(row=r, column=col_pas, value=round(s["pas"]))
        ws.cell(row=r, column=col_na, value=s["na"])
        # KET left blank

        for c in range(1, last_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER_ALL
            if c == 2:
                cell.font = FONT_NAMA
                cell.alignment = ALIGN_LEFT
            else:
                cell.font = FONT_BODY
                cell.alignment = ALIGN_CENTER

    # ---- footer ----
    footer_row = start_row + len(students) + 1
    ws.cell(row=footer_row + 1, column=col_pts, value="Sragen,\u2026....................")
    NA_formula = ws.cell(row=footer_row + 2, column=2, value="   NA = (3X NH)+PSTS+PSAS")
    NA_formula.font = Font(underline="single")
    NA_formula.border = Border(top=THIN, left=THIN, right=THIN)
    NA_denominator = ws.cell(row=footer_row + 3, column=2, value=5)
    NA_denominator.alignment = Alignment(horizontal="center")
    NA_denominator.border = Border(bottom=THIN, left=THIN, right=THIN)

    ws.cell(row=footer_row + 2, column=col_pts, value="Guru Mapel")
    ws.cell(row=footer_row + 5, column=col_pts, value=TEACHER_NAME)

    # ---- column widths ----
    ws.column_dimensions["A"].width = 5.14
    ws.column_dimensions["B"].width = 34.71
    for c in range(col_cp_start, col_cp_end + 1):
        ws.column_dimensions[L(c)].width = 4.57
    ws.column_dimensions[L(col_rata)].width = 12
    ws.column_dimensions[L(col_pts)].width = 8
    ws.column_dimensions[L(col_pas)].width = 8
    ws.column_dimensions[L(col_na)].width = 8
    ws.column_dimensions[L(col_ket)].width = 8.71

    return ws


def write_warnings_sheet(wb, warnings):
    ws = wb.create_sheet("Peringatan")
    ws["A1"] = "Peringatan (perlu dicek manual)"
    ws["A1"].font = Font(bold=True, size=12)
    ws.column_dimensions["A"].width = 90
    if not warnings:
        ws["A2"] = "Tidak ada peringatan."
        return
    for i, w in enumerate(warnings, start=2):
        ws.cell(row=i, column=1, value=w)


def build_output(classes, warnings, out_path):
    wb = Workbook()
    wb.remove(wb.active)
    for cname, cdata in classes.items():
        write_class_sheet(wb, cname, cdata)
    write_warnings_sheet(wb, warnings)
    wb.save(out_path)


# ----------------------------------------------------------------------

def main():
    if len(sys.argv) != 3:
        print("Usage: python na_report.py input.xlsx output.xlsx")
        sys.exit(1)
    in_path, out_path = sys.argv[1], sys.argv[2]
    classes, warnings = read_input(in_path)
    classes = compute(classes)
    build_output(classes, warnings, out_path)
    print(f"Selesai. {len(classes)} kelas diproses, {len(warnings)} peringatan.")
    if warnings:
        print("Lihat sheet 'Peringatan' di file output untuk detail.")


if __name__ == "__main__":
    main()
