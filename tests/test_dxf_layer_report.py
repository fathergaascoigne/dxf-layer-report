import math
import tempfile
import unittest
from pathlib import Path

import ezdxf
import openpyxl

from src.dxf_layer_measure import (
    ScanOptions,
    collect_dxf_tasks,
    entity_length,
    measure_layers_in_document,
    normalize_layers,
    scan_dxf_files,
    write_results,
)


class TestNormalizeLayers(unittest.TestCase):
    def test_deduplicates_case_insensitive_and_keeps_first_spelling(self):
        self.assertEqual(normalize_layers([" CUT ", "cut", "Cut", "A"]), ("CUT", "A"))


class TestMeasureLayers(unittest.TestCase):
    def test_matches_layers_case_insensitively(self):
        doc = ezdxf.new(setup=True)
        msp = doc.modelspace()
        msp.add_line((0, 0), (10, 0), dxfattribs={"layer": "cut"})

        measures, status = measure_layers_in_document(
            doc,
            target_layers=("CUT",),
            unit_to_mm_factor=1.0,
            spline_segments=96,
            ellipse_segments=96,
        )

        self.assertEqual(status, "OK")
        self.assertEqual(measures["CUT"], 10.0)

    def test_lwpolyline_bulge_arc_length_is_accounted_for(self):
        doc = ezdxf.new(setup=True)
        msp = doc.modelspace()
        bulge = math.tan(math.pi / 8.0)
        poly = msp.add_lwpolyline([(0, 0, bulge), (10, 0, 0)], format="xyb", dxfattribs={"layer": "ARC"})

        length = entity_length(poly, spline_segments=96, ellipse_segments=96)
        expected = (10.0 / math.sqrt(2.0)) * (math.pi / 2.0)
        self.assertAlmostEqual(length, expected, places=6)


class TestScanAndWorkbook(unittest.TestCase):
    def test_collect_tasks_counts_folders(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            sub = root / "sub"
            sub.mkdir()
            (root / "a.dxf").write_text("0\nEOF\n", encoding="utf-8")
            (sub / "b.dxf").write_text("0\nEOF\n", encoding="utf-8")

            tasks, folder_count = collect_dxf_tasks(root, recursive=True)

            self.assertEqual(folder_count, 2)
            self.assertEqual(len(tasks), 2)
            self.assertEqual(tasks[0].folder_rel, ".")
            self.assertEqual(tasks[1].folder_rel, "sub")

    def test_scan_filters_missing_layers_and_writes_numeric_cells(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)

            doc_ok = ezdxf.new(setup=True)
            doc_ok.modelspace().add_line((0, 0), (25, 0), dxfattribs={"layer": "14"})
            ok_path = root / "ok.dxf"
            doc_ok.saveas(ok_path)

            doc_missing = ezdxf.new(setup=True)
            doc_missing.modelspace().add_line((0, 0), (5, 0), dxfattribs={"layer": "OTHER"})
            missing_path = root / "missing.dxf"
            doc_missing.saveas(missing_path)

            options = ScanOptions(
                source_path=root,
                target_layers=("14",),
                recursive=False,
                workers=1,
                include_missing_layers=False,
                include_read_errors=False,
            )
            results, total_files, stopped_early = scan_dxf_files(options)

            self.assertFalse(stopped_early)
            self.assertEqual(total_files, 2)
            self.assertEqual(len(results), 1)
            self.assertEqual(results[0].filename, "ok.dxf")
            self.assertEqual(results[0].measures_mm["14"], 25.0)

            output = root / "out.xlsx"
            write_results(output, options.target_layers, results)
            workbook = openpyxl.load_workbook(output)
            sheet = workbook["Measurements"]

            self.assertEqual(sheet["A2"].value, "ok.dxf")
            self.assertEqual(sheet["B2"].value, 25.0)
            self.assertEqual(sheet["B2"].number_format, '0.000" mm"')


if __name__ == "__main__":
    unittest.main()

class TestThreadedConsistency(unittest.TestCase):
    def test_single_worker_and_multi_worker_return_same_measures(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            for idx in range(6):
                doc = ezdxf.new(setup=True)
                doc.modelspace().add_line((0, 0), (idx + 1, 0), dxfattribs={"layer": "CUT"})
                doc.saveas(root / f"file_{idx}.dxf")

            base_kwargs = dict(
                source_path=root,
                target_layers=("CUT",),
                recursive=False,
                include_missing_layers=True,
                include_read_errors=True,
            )
            sequential_results, _, _ = scan_dxf_files(ScanOptions(workers=1, **base_kwargs))
            threaded_results, _, _ = scan_dxf_files(ScanOptions(workers=4, **base_kwargs))

            seq_pairs = [(item.filename, item.measures_mm["CUT"], item.status) for item in sequential_results]
            thr_pairs = [(item.filename, item.measures_mm["CUT"], item.status) for item in threaded_results]
            self.assertEqual(seq_pairs, thr_pairs)
