import shutil
import tempfile
import unittest
import zipfile
from copy import deepcopy
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import yaml

from generate import LAYOUT, ROOT, build, convert_ods, load_config, verify_ods

CONFIG = ROOT / "programmes/digital-transformation-management.yml"
SOFFICE = shutil.which("libreoffice") or shutil.which("soffice")
if not SOFFICE and Path("/Applications/LibreOffice.app/Contents/MacOS/soffice").exists():
    SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"


class GeneratorTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.config = load_config(CONFIG)

    def write_config(self, config):
        path = self.root / "custom.yml"
        path.write_text(yaml.safe_dump(config, allow_unicode=True), encoding="utf-8")
        return path

    def test_preserve_all_formulas_and_cell_protection(self):
        base = openpyxl.load_workbook(LAYOUT)
        result = openpyxl.load_workbook(build(CONFIG, self.root))
        for sheet in base:
            if sheet.title.endswith("legend"):
                continue
            for row in sheet:
                for cell in row:
                    actual = result[sheet.title][cell.coordinate]
                    self.assertEqual(
                        (cell.protection.locked, cell.protection.hidden),
                        (actual.protection.locked, actual.protection.hidden),
                    )
                    if cell.data_type == "f":
                        self.assertEqual(cell.value, actual.value)

    def test_cluster_lists_can_grow_past_the_original_layout(self):
        config = deepcopy(self.config)
        config["degree_classes"].append(
            {
                "name": "Design",
                "description": "Design methods and visual communication.",
                "codes": [{"code": "EXAMPLE", "en": "Design", "it": "Design"}],
            }
        )
        workbook = openpyxl.load_workbook(build(self.write_config(config), self.root))
        self.assertEqual(workbook["Degree class legend"]["A11"].value, "Design")
        self.assertEqual(workbook["Degree class legend"]["A12"].value, "Other")
        dv = next(d for d in workbook["Degrees"].data_validations.dataValidation if "N10" in d)
        sheet, area = next(workbook.defined_names[dv.formula1].destinations)
        self.assertEqual([r[0].value for r in workbook[sheet][area]][-2:], ["Design", "Other"])

    def test_custom_programme_clusters_topics_colors_and_protection(self):
        cfg = deepcopy(self.config)
        cfg["programme"].update(
            slug="data-science",
            name="Data Science",
            topics=["machine learning"],
            deadline="2027-05-20",
        )
        cfg["degree_classes"] = [cfg["degree_classes"][1]]
        cfg["exam_sectors"] = [cfg["exam_sectors"][1]]
        cfg["colors"]["palette"].update(input="#E0F0D0", error_fill="#FFCCCC", border="#112233")
        cfg["colors"]["tabs"]["Degrees"] = "#AA5500"
        workbook = openpyxl.load_workbook(build(self.write_config(cfg), self.root))
        self.assertEqual(workbook["Instructions"]["A2"].value, "Data Science")
        self.assertEqual(workbook["Degrees"].sheet_properties.tabColor.rgb[-6:], "AA5500")
        self.assertEqual(workbook["Degrees"]["B10"].fill.fgColor.rgb[-6:], "E0F0D0")
        self.assertEqual(workbook["Degrees"]["B10"].border.left.color.rgb[-6:], "112233")
        self.assertEqual(
            workbook["Personal & contacts"]["B34"].value.date().isoformat(), "2027-05-20"
        )
        self.assertTrue(
            any("machine learning" in str(c.value) for c in workbook["Work experience"][8])
        )
        for name, column, expected in [
            ("Degrees", "N", "Statistics"),
            ("Exams", "Q", "Computer science"),
        ]:
            s = workbook[name]
            dv = next(d for d in s.data_validations.dataValidation if f"{column}10" in d)
            source, area = next(workbook.defined_names[dv.formula1].destinations)
            self.assertEqual([r[0].value for r in workbook[source][area]], [expected, "Other"])
            self.assertTrue(s["A10"].protection.locked)
            self.assertFalse(s["B10"].protection.locked)
        self.assertEqual(workbook["Degree class legend"]["A7"].value, "Statistics")
        self.assertEqual(workbook["Degree class legend"]["A8"].value, "Other")
        for name in ("Degree class legend", "Exam sector legend"):
            self.assertFalse(workbook[name].tables)
            self.assertIsNone(workbook[name].auto_filter.ref)
        self.assertTrue(all(s.protection.sheet for s in workbook))
        self.assertEqual(workbook["Lists"].sheet_state, "hidden")
        self.assertTrue(workbook["Exams"]["N11"].value.startswith('=IF(OR($B11="",$C11<>$C10,'))
        self.assertFalse(workbook["Exams"]["N11"].protection.locked)
        self.assertTrue(workbook["Exams"]["G11"].protection.locked)
        self.assertEqual(workbook["Personal & contacts"]["B9"].number_format, "yyyy-mm-dd")
        rules = next(iter(workbook["Degrees"].conditional_formatting._cf_rules.values()))
        self.assertEqual(rules[0].dxf.fill.bgColor.rgb[-6:], "FFCCCC")

    def test_reject_ambiguous_or_unsafe_configuration(self):
        variants = []
        bad = deepcopy(self.config)
        bad["degree_classes"][1]["codes"] = bad["degree_classes"][0]["codes"]
        variants.append(bad)
        bad = deepcopy(self.config)
        bad["exam_sectors"][0]["name"] = "Other"
        variants.append(bad)
        bad = deepcopy(self.config)
        bad["colors"]["palette"]["input"] = "red"
        variants.append(bad)
        bad = deepcopy(self.config)
        bad["programme"]["slug"] = "../outside"
        variants.append(bad)
        bad = deepcopy(self.config)
        bad["programme"]["name"] = "=SUM(1,2)"
        variants.append(bad)
        bad = deepcopy(self.config)
        bad["colours"] = {}
        variants.append(bad)
        for config in variants:
            with self.subTest(config=config), self.assertRaises(ValueError):
                load_config(self.write_config(config))

    @unittest.skipUnless(SOFFICE, "LibreOffice not installed")
    def test_ods_retains_protection_validation_and_recalculates_defaults(self):
        xlsx = build(CONFIG, self.root)
        workbook = openpyxl.load_workbook(xlsx)
        s = workbook["Exams"]
        s["B10"], s["C10"], s["N10"], s["B11"] = "Algorithms", "D001", 30, "Networks"
        workbook.save(xlsx)
        ods = convert_ods(xlsx, SOFFICE)
        verify_ods(ods)
        ns = {
            "t": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
            "o": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
        }
        with zipfile.ZipFile(ods) as archive:
            root = ET.fromstring(archive.read("content.xml"))
        exams = next(
            t for t in root.findall(".//t:table", ns) if t.get(f"{{{ns['t']}}}name") == "Exams"
        )
        row_index = 0
        value = None
        for row in exams.findall("t:table-row", ns):
            repeats = int(row.get(f"{{{ns['t']}}}number-rows-repeated", "1"))
            if row_index < 11 <= row_index + repeats:
                col_index = 0
                for cell in row:
                    count = int(cell.get(f"{{{ns['t']}}}number-columns-repeated", "1"))
                    if col_index < 14 <= col_index + count:
                        value = cell.get(f"{{{ns['o']}}}value")
                        break
                    col_index += count
                break
            row_index += repeats
        self.assertEqual(float(value), 30)


if __name__ == "__main__":
    unittest.main()
