# Nilai NA — grade calculator + site CSV filler

Scripts:

1. **`na_report.py`** reads a class-by-class grade sheet and produces the
   official "Pengolahan Nilai Rapor" workbook, with Rata-rata NH and NA
   calculated per student.
2. **`fill_site_csv.py`** takes that report and fills the online
   submission site's per-class CSV templates (matching students by name).

## Requirements

- Python 3.9+
- [openpyxl](https://openpyxl.readthedocs.io/) — `pip3 install openpyxl`

## 1. Generate the NA report

```
python na_report.py input.xlsx output.xlsx

#Change "input.xlsx" to whatever the workbook filename is
#Change "output.xlsx" to whatever name you want the output to be
```

**Input** (`input.xlsx`): one sheet per class, header row exactly:

| Nama | Tugas dan Latihan | Nilai 1 | ... | Nilai 8 | Ulangan Harian | PTS | PAS |
|---|---|---|---|---|---|---|---|

- `Nilai 1`–`Nilai 8`: number 0–100, or a letter grade (`A`, `B`, `B+`, or
  combinations like `AA` = A + A). Numbers are averaged; letters become
  extra points (A = 3, B+ = 2, B = 1).
- `Tugas dan Latihan` and `Ulangan Harian` are optional — leave blank if a
  class doesn't use them.
- Blank rows and blank PTS/PAS are handled but flagged in the warnings.

**Output** (`output.xlsx`): one sheet per class in the school's official
layout (dynamic number of CP columns, min. 6), plus a `Peringatan` sheet
listing anything that needed a default (unrecognized grade, missing
PTS/PAS, students with no harian data at all).

**Calculation, per student:**

- `Rata-rata NH` = average of numeric Nilai 1–8 **and Ulangan Harian** (if
  the class uses it) + letter-grade points + tugas dan latihan points.
  Shown in the output only.
- `Nilai Harian` (used inside NA) = average of numeric Nilai 1–8 **only**
  (UH excluded) + letter-grade points + tugas dan latihan points.
- Class **without** UH: `NA = Nilai Harian×40% + PTS×30% + PAS×30%`
- Class **with** UH: `NA = Nilai Harian×20% + UH×35% + PTS×20% + PAS×25%`
- Both `Rata-rata NH` and `NA` are rounded to whole numbers.

All of the above (weights, letter→point mapping, min CP columns, header
text) live at the top of `na_report.py` under `CONFIG` — edit there, not
in the logic below it.

## 2. Fill the site's CSV templates

```
python fill_site_csv.py output.xlsx templates_dir/ result_dir/

#Change "output.xlsx" to whatever the "na_report.py" output filename is
#Change "templates_dir/" and "result_dir/" to whatever the folder names are
```

- `templates_dir/` — the site's `Template_Nilai_*.csv` files, one per
  class, each with `NISN, Nama, Nilai` columns (NISN pre-filled, Nilai
  blank).
- Matches each template to a class sheet in `output.xlsx` by name, then
  matches students by `Nama` (case/whitespace-insensitive).
- Writes filled CSVs to `result_dir/` with the `Template_` prefix
  stripped, e.g. `Template_Nilai_PTS_X TJKT 1.csv` → `Nilai_PTS_X TJKT 1.csv`.
- Any student on the site but not in the report (or vice versa) is left
  blank / skipped and logged in `result_dir/Peringatan_Pencocokan.txt`
  instead of guessed at.

## Known limitations / things to revisit

- CP column count in the output is based on how many `Nilai N` headers
  exist in the input sheet (currently 8), not how many a class actually
  fills in.
- Name matching for the CSV filler is exact after normalizing
  case/whitespace — nicknames or differently-spelled names won't match
  automatically.
- School name, subject, semester, and teacher name are left blank in the
  output layout; fill in `CONFIG` at the top of `na_report.py` if you want
  them hardcoded.
