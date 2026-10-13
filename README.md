# tapFill

Generate one filled TAP PDF for each nonempty personnel row in an Excel workbook.
Python 3.10 or newer is required. The supplied `TAP_fillable.pdf` is included.

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python fill_tap.py personnel.xlsx --output-dir filled_pdfs
```

The first worksheet is used by default. Select another with `--sheet "Roster"`,
or use another copy of the same form with `--template path/to/TAP_fillable.pdf`.
The first nonempty row must contain column headers; blank rows are skipped.

| XLSX column | PDF field |
| --- | --- |
| Name: Last, First M | Last name, first name, middle initial |
| Rank | Rank and corresponding enlisted grade E1–E9 |
| Contact Number | Cell phone |
| Personal Email | Work email (blank when Personal Email is missing) |
| DOB | Date of birth, formatted MM/DD/YYYY |
| DOD ID | DoD ID |
| Home of Record | Post-transition location |
| Gender | Gender |

Component is always **RC**, UIC is **WS64AA**, and service affiliation is **Army**.
Grades: PVT/PV1 → E1, PV2 → E2, PFC → E3, SPC/CPL → E4, SGT → E5,
SSG → E6, SFC → E7, MSG/1SG → E8, and SGM/CSM/SMA → E9. Full rank names
and E1–E9 (including hyphens) are also accepted. Unknown ranks leave grade blank.

Other roster columns are ignored because the TAP form has no corresponding
fields. Missing values remain blank, including the yes/no checkboxes. Existing
printed form defaults, including the TAP center and leaving reason, are retained.
Army Email is ignored; the PDF's Work Email field uses Personal Email only.

Optional extra columns: `Separation Date` (or `REFRAD`), `Departure Date`, `MOS`,
`Unit Name`, `Salary`, and `Field of Interest` (or `Job Title`). When a separation
date is supplied without a departure date, it fills both dates. The anniversary
year ending date is not used as a separation date.

Names must use `Last, First M`, for example `Smith, Jane A`. A trailing single
letter or letter with a period is treated as the middle initial. Without a comma,
the entire name goes in the last-name field; the script does not guess name order.
Store DoD IDs and phone numbers as Excel text to preserve leading zeros. Simple
Excel zero-padding formats and native Excel dates are also supported. Formulas
use Excel's last saved calculated values; formulas without cached values are blank.

Outputs are named `Smith_tapForm.pdf`. Duplicate surnames or existing files get
`Smith_2_tapForm.pdf`, `Smith_3_tapForm.pdf`, etc., to preserve every PDF. Missing
names use `unnamed_row_2_tapForm.pdf` with the actual Excel row number. Unsafe
filename characters are replaced with underscores. PDFs remain editable and
include field appearances so filled values display in PDF viewers.

Run the tests with:

```sh
.venv/bin/python -m unittest discover -s tests -v
```
