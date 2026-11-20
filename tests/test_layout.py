import subprocess
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import openpyxl
from layout_snapshot import diff, snapshot
from test_generate import SOFFICE

from generate import LAYOUT
from tools.build_layout import build


def headers(ws, row=9):
    return {str(c.value).removesuffix(" *"): c.column_letter for c in ws[row] if c.value}


class LayoutTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)

    def test_builder_reproduces_committed_layout(self):
        # If this fails after an intended change, rebuild: uv run poe layout
        built = build(self.root / "layout.xlsx")
        self.assertEqual(diff(snapshot(LAYOUT), snapshot(built)), [])

    @unittest.skipUnless(SOFFICE, "LibreOffice not installed")
    def test_formulas_recalculate_as_designed(self):
        # Port of the original builder's self-checks, evaluated by LibreOffice.
        def scenario(name, override):
            wb = openpyxl.load_workbook(LAYOUT)
            deg, ex, work = wb["Degrees"], wb["Exams"], wb["Work experience"]
            d, e = headers(deg), headers(ex)
            deg["B10"], deg["H10"] = "Test degree", "Graduated"
            deg["G10"], deg["I10"] = datetime(2021, 9, 1), datetime(2024, 9, 1)
            deg[d["Degree cluster"] + "10"] = "Computing & information engineering"
            deg[d["Credit system"] + "10"], deg[d["Hours per credit"] + "10"] = "Local credits", 25
            deg[d["Final scale minimum"] + "10"], deg[d["Final scale maximum"] + "10"] = 0, 110
            deg["B11"], deg["B12"] = "Second degree", "Third degree"
            ex["B10"], ex["C10"] = "Test exam", "D001"
            ex[e["Exam cluster"] + "10"] = "Computer science"
            ex[e["Scale minimum"] + "10"], ex[e["Scale maximum"] + "10"] = 0, 30
            ex["B11"] = "Second exam"
            wb["Personal & contacts"]["B34"] = datetime(2026, 11, 20)
            work["B10"], work["F10"] = "Test role", datetime(2026, 1, 1)
            work["G10"], work["H10"] = datetime(2026, 3, 31), "No"
            if override:
                deg[d["Final scale maximum"] + "11"] = 100
                ex["C11"] = "D002"
                work["G10"], work["H10"] = None, "Yes"
            path = self.root / f"{name}.xlsx"
            wb.save(path)
            return path, d, e

        base, d, e = scenario("base", override=False)
        changed, _, _ = scenario("changed", override=True)
        out = self.root / "out"
        subprocess.run(
            [
                str(SOFFICE),
                f"-env:UserInstallation={(self.root / 'profile').as_uri()}",
                "--headless",
                "--convert-to",
                "xlsx",
                "--outdir",
                str(out),
                str(base),
                str(changed),
            ],
            check=True,
            capture_output=True,
            timeout=300,
        )
        wb = openpyxl.load_workbook(out / "base.xlsx", data_only=True)
        deg, ex, work = wb["Degrees"], wb["Exams"], wb["Work experience"]
        self.assertEqual(deg["A10"].value, "D001")
        self.assertEqual(ex["A10"].value, "EX001")
        self.assertAlmostEqual(deg["L10"].value, 36, delta=0.2)
        self.assertEqual(ex["G10"].value, "Local credits")
        self.assertEqual(ex["H10"].value, 25)
        self.assertNotIn("Unknown degree", str(ex[e["Check"] + "10"].value))
        self.assertNotIn("Invalid choice: Degree cluster", str(deg[d["Check"] + "10"].value))
        # Defaults stay blank in unused rows, preserve zero and propagate.
        self.assertIn(deg[d["Final scale maximum"] + "13"].value, (None, ""))
        self.assertEqual(deg[d["Final scale minimum"] + "11"].value, 0)
        self.assertEqual(deg[d["Final scale maximum"] + "12"].value, 110)
        self.assertEqual(ex["C11"].value, "D001")
        self.assertEqual(ex[e["Scale maximum"] + "11"].value, 30)
        self.assertAlmostEqual(work["I10"].value, 3, delta=0.1)

        wb = openpyxl.load_workbook(out / "changed.xlsx", data_only=True)
        deg, ex, work = wb["Degrees"], wb["Exams"], wb["Work experience"]
        self.assertEqual(deg[d["Final scale maximum"] + "12"].value, 100)  # override propagates
        self.assertIn(ex[e["Scale maximum"] + "11"].value, (None, ""))  # degree change resets
        self.assertGreater(work["I10"].value, 10)  # ongoing: runs to reference date


if __name__ == "__main__":
    unittest.main()
