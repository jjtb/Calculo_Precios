from __future__ import annotations

from datetime import datetime
from dataclasses import replace
from io import BytesIO
import math
from pathlib import Path
import tempfile
import unittest

import pyarrow.parquet as pq
from openpyxl import load_workbook

from charts import marker_chart, oil_marker_series
from pricing import (
    CATALOG_PATH,
    OIL_REFERENCE_FIELDS,
    PROPERTY_PATH,
    REPORT_END_YEAR,
    REPORT_START_YEAR,
    calculate_prices,
    load_catalog,
    load_markers,
    load_properties,
    oil_price,
    validate_properties,
)
from storage import write_calculation


ROOT = Path(__file__).resolve().parents[1]
DELIVERABLES = ROOT / "Data_Informacion" / "Entregables" / "EntregablesCartera-2025"
# Historical validation values; final reports use the current markers for Nejo.
VALIDATION_REFERENCE_EXCEPTIONS = {
    ("Nejo", "Condensado"),
    ("Nejo", "Gas"),
}


class PricingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.properties, cls.duplicate_issues = load_properties(PROPERTY_PATH)
        cls.catalog, _ = load_catalog(CATALOG_PATH)

    def test_api_interpolation_and_reference_markers(self):
        markers = load_markers("alto")
        oil = oil_price(31, markers, 2025)
        gas = 1366.4588701969253 * markers.gas_usd_per_mmbtu["GH Amargo Sureste"][2025] / 1000
        condensate = markers.condensate_usd_per_bbl["Condensado Amargo"][2025]
        gas_bn = markers.gas_bn_usd_per_mpc[2025]
        self.assertAlmostEqual(oil, 96.76745680187089, places=10)
        self.assertAlmostEqual(gas, 6.302169153819367, places=10)
        self.assertAlmostEqual(condensate, 43.1799861987279, places=10)
        self.assertAlmostEqual(gas_bn, 6.913306805840032, places=10)

    def test_oil_chart_includes_annual_homol_ket_and_kayab_references(self):
        markers = load_markers("medio")
        series = oil_marker_series(markers)
        expected_2025 = {
            "Homol (37.1° API)": 67.3475914522,
            "Ket (15.37° API)": 56.7713142715,
            "Kayab (8.6° API)": 52.2890880024,
        }
        for name, value in expected_2025.items():
            with self.subTest(reference=name):
                self.assertAlmostEqual(series[name][0], value, places=8)
                self.assertEqual(len(series[name]), len(markers.years))
                self.assertAlmostEqual(
                    series[name][1],
                    oil_price(
                        OIL_REFERENCE_FIELDS[name.split(" (")[0]],
                        markers,
                        markers.years[1],
                    ),
                )
                self.assertNotEqual(series[name][0], series[name][1])

        svg = marker_chart("Marcadores de aceite", "USD/b", markers.years, series)
        self.assertIn("Homol (37.1° API)", svg)
        self.assertIn("Ket (15.37° API)", svg)
        self.assertIn("Kayab (8.6° API)", svg)
        self.assertIn('viewBox="0 0 720 277"', svg)

    def test_nejo_final_report_uses_current_model_markers(self):
        nejo = {
            prop.hydrocarbon: prop
            for prop in self.properties
            if prop.field == "Nejo"
        }
        for scenario in ("alto", "bajo", "medio"):
            with self.subTest(scenario=scenario):
                markers = load_markers(scenario)
                rows = calculate_prices(
                    self.properties,
                    self.catalog,
                    markers,
                    calculated_at="2026-11-14T12:00:00-06:00",
                    run_id="reference",
                    calculation_mode="reporte_final",
                )
                result_map = {
                    (row["campo"], row["tipo_hidrocarburo"]): row for row in rows
                }
                condensate = result_map[("Nejo", "Condensado")]
                gas = result_map[("Nejo", "Gas")]
                self.assertAlmostEqual(
                    condensate["AÑO 2025"],
                    markers.condensate_usd_per_bbl["Condensado Dulce"][2025],
                )
                self.assertAlmostEqual(
                    gas["AÑO 2025"],
                    nejo["Gas"].api_btu
                    * markers.gas_usd_per_mmbtu["GH Dulce Norte"][2025]
                    / 1000,
                )

    def test_nejo_final_report_overrides_oil_flag_for_condensate(self):
        nejo_condensate = next(
            prop
            for prop in self.properties
            if prop.field == "Nejo" and prop.hydrocarbon == "Condensado"
        )
        flagged_properties = [
            replace(prop, calculated_as_oil=13)
            if prop is nejo_condensate
            else prop
            for prop in self.properties
        ]
        markers = load_markers("alto")
        validation_rows = calculate_prices(
            flagged_properties,
            self.catalog,
            markers,
            calculated_at="2026-11-14T12:00:00-06:00",
            run_id="validation",
        )
        final_rows = calculate_prices(
            flagged_properties,
            self.catalog,
            markers,
            calculated_at="2026-11-14T12:00:00-06:00",
            run_id="final",
            calculation_mode="reporte_final",
        )
        validation_price = next(
            row["AÑO 2025"]
            for row in validation_rows
            if row["campo"] == "Nejo" and row["tipo_hidrocarburo"] == "Condensado"
        )
        final_price = next(
            row["AÑO 2025"]
            for row in final_rows
            if row["campo"] == "Nejo" and row["tipo_hidrocarburo"] == "Condensado"
        )
        self.assertAlmostEqual(
            validation_price,
            oil_price(nejo_condensate.api_btu, markers, 2025),
        )
        self.assertAlmostEqual(
            final_price,
            markers.condensate_usd_per_bbl["Condensado Dulce"][2025],
        )

    def test_all_three_scenarios_match_2025_deliverables(self):
        for scenario in ("alto", "bajo", "medio"):
            with self.subTest(scenario=scenario):
                markers = load_markers(scenario)
                errors, _ = validate_properties(
                    self.properties, self.catalog, self.duplicate_issues
                )
                self.assertEqual(errors, [])
                rows = calculate_prices(
                    self.properties,
                    self.catalog,
                    markers,
                    calculated_at="2026-11-14T12:00:00-06:00",
                    run_id="reference",
                )
                result_map = {
                    (row["campo"], row["tipo_hidrocarburo"]): row for row in rows
                }
                reference_path = next(
                    DELIVERABLES.glob(f"*{scenario.title()}*.xlsx")
                )
                workbook = load_workbook(
                    reference_path, read_only=True, data_only=True
                )
                try:
                    sheet = workbook.worksheets[0]
                    expected_count = sheet.max_row - 1
                    self.assertEqual(len(rows), expected_count)
                    for expected in sheet.iter_rows(min_row=2, values_only=True):
                        field, hydrocarbon = expected[4], expected[7]
                        if (field, hydrocarbon) in VALIDATION_REFERENCE_EXCEPTIONS:
                            continue
                        actual = result_map[(field, hydrocarbon)]
                        self.assertEqual(actual["tipo_gas"], expected[5])
                        self.assertEqual(actual["unidad"], expected[8])
                        self.assertAlmostEqual(
                            float(actual["api_btu"]), float(expected[6]), places=10
                        )
                        for offset, year in enumerate(
                            range(REPORT_START_YEAR, REPORT_END_YEAR + 1), start=9
                        ):
                            value = expected[offset]
                            if value is None:
                                continue
                            self.assertTrue(
                                math.isclose(
                                    float(actual[f"AÑO {year}"]),
                                    float(value),
                                    rel_tol=1e-11,
                                    abs_tol=1e-9,
                                ),
                                f"{scenario} / {field} / {hydrocarbon} / {year}: "
                                f"{actual[f'AÑO {year}']} != {value}",
                            )
                finally:
                    workbook.close()

    def test_duplicate_bamoa_rows_use_first_reference_record(self):
        bamoa = [
            prop
            for prop in self.properties
            if prop.field == "Bamoa" and prop.hydrocarbon == "Aceite"
        ]
        self.assertEqual(len(bamoa), 1)
        self.assertAlmostEqual(bamoa[0].api_btu, 24.731010685824483)
        self.assertTrue(
            any("Bamoa / Aceite" in issue for issue in self.duplicate_issues)
        )

    def test_calculation_is_saved_as_parquet_with_metadata(self):
        markers = load_markers("medio")
        run_id = "abcdef12-1234"
        rows = calculate_prices(
            self.properties,
            self.catalog,
            markers,
            calculated_at=datetime.now().astimezone().isoformat(timespec="seconds"),
            run_id=run_id,
            calculation_mode="reporte_final",
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            from storage import RESULTS_DIR
            import storage

            original = storage.RESULTS_DIR
            storage.RESULTS_DIR = Path(temp_dir)
            try:
                output = write_calculation(
                    rows[:4],
                    run_id=run_id,
                    scenario="Medio",
                    calculated_at="2026-11-14T12:00:00-06:00",
                    calculation_mode="reporte_final",
                    input_set_version=3,
                    source_files=["propiedades.xlsx"],
                    warnings=[],
                )
                table = pq.read_table(output)
                self.assertEqual(table.num_rows, 4)
                self.assertEqual(
                    table.schema.metadata[b"process"].decode("utf-8"), "Cartera"
                )
                self.assertEqual(
                    table.schema.metadata[b"calculation_mode"].decode("utf-8"),
                    "reporte_final",
                )
                self.assertEqual(table.schema.metadata[b"input_set_version"], b"3")
                self.assertEqual(
                    table.column("modo_calculo")[0].as_py(), "reporte_final"
                )
                self.assertEqual(table.column("version_insumos")[0].as_py(), 3)
                self.assertTrue((Path(temp_dir) / "bitacora.parquet").is_file())
                excel_bytes, filename = storage.calculation_xlsx(run_id)
                workbook = load_workbook(BytesIO(excel_bytes), data_only=True)
                try:
                    self.assertEqual(
                        filename,
                        "precios_cartera_medio_"
                        + datetime.fromisoformat(
                            str(rows[0]["fecha_calculo"])
                        ).strftime("%Y%m%d")
                        + "_version_03.xlsx",
                    )
                    self.assertEqual(
                        workbook["Precios Cartera"].max_column,
                        105,
                    )
                    self.assertEqual(
                        workbook["Precios Cartera"].cell(1, 6).value,
                        "AÑO 2025",
                    )
                    self.assertEqual(
                        workbook["Precios Cartera"].cell(1, 105).value,
                        f"AÑO {REPORT_END_YEAR}",
                    )
                    self.assertEqual(
                        workbook["Metadatos"]["B6"].value,
                        "reporte_final",
                    )
                finally:
                    workbook.close()
                log_path = Path(temp_dir) / "bitacora.parquet"
                legacy_log = pq.read_table(log_path).drop(
                    ["modo_calculo", "version_insumos"]
                )
                pq.write_table(legacy_log, log_path)
                write_calculation(
                    rows[:2],
                    run_id="test-run-2",
                    scenario="Medio",
                    calculated_at="2026-11-14T12:01:00-06:00",
                    calculation_mode="validacion",
                    input_set_version=4,
                    source_files=["propiedades.xlsx"],
                    warnings=[],
                )
                log = pq.read_table(log_path)
                self.assertEqual(log.num_rows, 2)
                self.assertIsNone(log.column("modo_calculo")[0].as_py())
                self.assertEqual(
                    log.column("modo_calculo")[1].as_py(), "validacion"
                )
                self.assertIsNone(log.column("version_insumos")[0].as_py())
                self.assertEqual(log.column("version_insumos")[1].as_py(), 4)
            finally:
                storage.RESULTS_DIR = original


if __name__ == "__main__":
    unittest.main()
