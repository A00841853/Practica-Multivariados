import contextlib
import csv
import io
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
import zipfile

from amd_monte_carlo import main, simulate
from amd_report import export_reports

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}


class ReportTests(unittest.TestCase):
    def test_csv_and_excel_agree_with_simulation(self):
        rows = simulate(604.72, simulations=1000)
        with tempfile.TemporaryDirectory() as folder:
            csv_path, excel_path = Path(folder) / "r.csv", Path(folder) / "r.xlsx"
            export_reports(rows, csv_path, excel_path)
            with csv_path.open(encoding="utf-8", newline="") as handle:
                actual = list(csv.DictReader(handle))
            self.assertEqual(len(actual), 16)
            self.assertEqual(list(actual[0]), list(rows[0]))
            for source, stored in zip(rows, actual):
                for key, value in source.items():
                    self.assertEqual(float(stored[key]), value)
            with zipfile.ZipFile(excel_path) as archive:
                charts = [p for p in archive.namelist()
                          if p.startswith("xl/charts/chart") and p.endswith(".xml")]
                self.assertEqual(len(charts), 4)
                self.assertEqual(len([p for p in archive.namelist()
                                      if p.startswith("xl/tables/table")]), 2)
                sheet = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
                cells = {c.attrib["r"]: c for c in sheet.findall(".//s:c", NS)}
                # Every raw numeric value equals the simulation, not a rounded display.
                for i, source in enumerate(rows, 8):
                    for j, value in enumerate(source.values()):
                        cell = cells[f"{chr(65+j)}{i}"]
                        self.assertAlmostEqual(float(cell.find("s:v", NS).text), value, places=9)
                summary = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
                cell = summary.find(".//s:c[@r='C9']", NS)
                self.assertEqual(cell.find("s:f", NS).text, "'Datos completos'!G8")
                self.assertAlmostEqual(float(cell.find("s:v", NS).text), rows[0]["mean"])
                self.assertEqual(summary.find(".//s:pane", NS).attrib["state"], "frozen")
                for name in charts:
                    chart = ET.fromstring(archive.read(name))
                    self.assertTrue(chart.findall(".//c:f", NS))
                    self.assertTrue(chart.findall(".//c:numCache/c:pt", NS))

    def test_default_cli_creates_both_files(self):
        old = Path.cwd()
        with tempfile.TemporaryDirectory() as folder:
            try:
                os.chdir(folder)
                with contextlib.redirect_stdout(io.StringIO()):
                    self.assertEqual(main(["--s0", "600", "--simulations", "10"]), 0)
                self.assertTrue(Path("resumen.csv").is_file())
                self.assertTrue(Path("analisis_amd.xlsx").is_file())
            finally:
                os.chdir(old)

    def test_custom_horizons_and_zero_volatility(self):
        rows = simulate(500, volatility=0, returns=(0.2,), months=(9, 2), simulations=1)
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)
            export_reports(rows, path / "custom.csv", path / "custom.xlsx")
            with zipfile.ZipFile(path / "custom.xlsx") as archive:
                chart = ET.fromstring(archive.read("xl/charts/chart3.xml"))
                values = chart.findall(".//c:xVal/c:numRef/c:numCache/c:pt/c:v", NS)
                self.assertEqual([float(v.text) for v in values], [2, 9])

    def test_missing_dependency_does_not_replace_files(self):
        rows = simulate(600, simulations=1)
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)
            with patch.dict("sys.modules", {"xlsxwriter": None}):
                with self.assertRaisesRegex(RuntimeError, "pip install"):
                    export_reports(rows, path / "r.csv", path / "r.xlsx")
            self.assertEqual(list(path.iterdir()), [])

    def test_invalid_extension_fails(self):
        with contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit) as caught:
                main(["--s0", "600", "--csv", "wrong.xlsx"])
        self.assertEqual(caught.exception.code, 2)


if __name__ == "__main__":
    unittest.main()
