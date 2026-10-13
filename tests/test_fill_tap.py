import contextlib
import io
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from pypdf import PdfReader

import fill_tap


TEMPLATE = Path(fill_tap.__file__).with_name("TAP_fillable.pdf")
HEADERS = [
    "", "Rank", "Name: Last, First M", "Contact Number", "&#x20;Army Email",
    "Personal Email", "DOB", "DOD ID", "Anniversary Year Ending Date (MM/DD)",
    "GTC/Last 4/ EXP", "CAC Exp:", "Home of Record", "ACFT", "H/W", "PSIP",
    "Food Restrictions/MRE Type", "Blood type", "Vaccine", "Deployment Orders Issued",
    "State Drivers License", "Military Drivers License", "DTS Login",
    "Gender<br><br><br>",
]


class FillTapTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def workbook(self, records):
        book = Workbook()
        sheet = book.active
        sheet.title = "Roster"
        sheet.append(HEADERS)
        for record in records:
            sheet.append([record.get(header) for header in HEADERS])
        path = self.root / "roster.xlsx"
        book.save(path)
        book.close()
        return path

    def run_batch(self, path, directory=None):
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            return fill_tap.fill_workbook(path, TEMPLATE, directory or self.root / "pdfs")

    def test_full_roster_fields_and_appearances(self):
        path = self.workbook([{
            "Rank": "SSG", "Name: Last, First M": "Smith, Jane A.",
            "Contact Number": "0123456789", "&#x20;Army Email": "jane@example.test",
            "Personal Email": "personal@example.test", "DOB": datetime(1990, 3, 4),
            "DOD ID": "0012345678", "Home of Record": "Texas",
            "Gender<br><br><br>": "Female", "Anniversary Year Ending Date (MM/DD)": "05/01",
        }])
        self.assertEqual(self.run_batch(path), (1, 0))
        reader = PdfReader(self.root / "pdfs" / "Smith_tapForm.pdf")
        values = {name: field.get("/V") for name, field in reader.get_fields().items()}
        expected = {
            "Last_Name": "Smith", "First_Name": "Jane", "Middle_Initial": "A",
            "Rank": "SSG", "Grade": "E6", "Component": "RC", "Unit_Code_UIC": "WS64AA",
            "Service_Affiliation": "Army", "Cell_Phone": "0123456789",
            "Work_Email": "personal@example.test", "Date_of_Birth": "03/04/1990",
            "DoD_ID": "0012345678", "Gender": "Female", "Post_Transition_Location": "Texas",
            "Separation_Date": "", "Departure_Date": "", "SkillBridge_Yes": "/Off",
            "Caregiver_No": "/Off",
        }
        for name, value in expected.items():
            self.assertEqual(values[name], value, name)
        last_widget = next(a.get_object() for a in reader.pages[0]["/Annots"]
                           if a.get_object().get("/T") == "Last_Name")
        self.assertIn(b"Smith", last_widget["/AP"]["/N"].get_data())
        self.assertIs(reader.trailer["/Root"]["/AcroForm"]["/NeedAppearances"].value, False)

    def test_missing_data_duplicate_surnames_and_rerun(self):
        path = self.workbook([
            {"Name: Last, First M": "Smith, Jane", "Rank": "PVT",
             "Personal Email": "fallback@example.test"},
            {"Name: Last, First M": "Smith, John", "Rank": "not a rank"},
            {},
            {"Rank": "SGT"},
        ])
        self.assertEqual(self.run_batch(path), (3, 0))
        directory = self.root / "pdfs"
        self.assertEqual({p.name for p in directory.iterdir()}, {
            "Smith_tapForm.pdf", "Smith_2_tapForm.pdf", "unnamed_row_5_tapForm.pdf",
        })
        fields = PdfReader(directory / "Smith_tapForm.pdf").get_fields()
        self.assertEqual(fields["Work_Email"]["/V"], "fallback@example.test")
        self.assertEqual(fields["Middle_Initial"]["/V"], "")
        self.assertEqual(fields["Date_of_Birth"]["/V"], "")
        self.assertEqual(PdfReader(directory / "Smith_2_tapForm.pdf").get_fields()["Grade"]["/V"], "")
        original = (directory / "Smith_tapForm.pdf").read_bytes()
        self.assertEqual(self.run_batch(path), (3, 0))
        self.assertEqual((directory / "Smith_tapForm.pdf").read_bytes(), original)
        self.assertEqual(len(list(directory.glob("*.pdf"))), 6)

    def test_all_standard_enlisted_ranks(self):
        expected = {"PVT": "E1", "PV1": "E1", "PV2": "E2", "PFC": "E3",
                    "SPC": "E4", "CPL": "E4", "SGT": "E5", "SSG": "E6",
                    "SFC": "E7", "MSG": "E8", "1SG": "E8", "SGM": "E9",
                    "CSM": "E9", "SMA": "E9", "staff sergeant": "E6", "E-9": "E9"}
        for rank, grade in expected.items():
            with self.subTest(rank=rank):
                self.assertEqual(fill_tap.grade_for_rank(rank), grade)
        self.assertEqual(fill_tap.grade_for_rank("CPT"), "")

    def test_missing_personal_email_does_not_use_army_email(self):
        path = self.workbook([{
            "Name: Last, First M": "Smith, Jane", "&#x20;Army Email": "army@example.test",
        }])
        self.assertEqual(self.run_batch(path), (1, 0))
        fields = PdfReader(self.root / "pdfs" / "Smith_tapForm.pdf").get_fields()
        self.assertEqual(fields["Work_Email"]["/V"], "")

    def test_optional_dates_numeric_padding_and_filename_safety(self):
        book = Workbook()
        sheet = book.active
        sheet.append(["Name: Last, First M", "DOD ID", "DOB", "Separation Date", "MOS"])
        sheet.append(["../Jones, Alex B", 12345, 32876, "2026-10-13", "25B"])
        sheet["B2"].number_format = "0000000000"
        path = self.root / "optional.xlsx"
        book.save(path)
        book.close()
        self.assertEqual(self.run_batch(path), (1, 0))
        files = list((self.root / "pdfs").glob("*.pdf"))
        self.assertEqual(len(files), 1)
        self.assertEqual(files[0].parent, self.root / "pdfs")
        fields = PdfReader(files[0]).get_fields()
        self.assertEqual(fields["DoD_ID"]["/V"], "0000012345")
        self.assertEqual(fields["Date_of_Birth"]["/V"], "01/03/1990")
        self.assertEqual(fields["Separation_Date"]["/V"], "10/13/2026")
        self.assertEqual(fields["Departure_Date"]["/V"], "10/13/2026")
        self.assertEqual(fields["MOS"]["/V"], "25B")

    def test_cli_sheet_and_errors(self):
        path = self.workbook([{"Rank": "SGT", "Name: Last, First M": "Lee, Pat"}])
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            self.assertEqual(fill_tap.main([str(path), "--sheet", "Roster", "--output-dir", str(self.root / "out")]), 0)
            self.assertEqual(fill_tap.main([str(path), "--sheet", "Missing"]), 1)
            self.assertEqual(fill_tap.main([str(self.root / "missing.xlsx")]), 1)

    def test_empty_roster_is_reported(self):
        path = self.workbook([])
        with self.assertRaisesRegex(ValueError, "no personnel rows"):
            self.run_batch(path)


if __name__ == "__main__":
    unittest.main()
