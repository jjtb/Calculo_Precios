from __future__ import annotations

import asyncio
from datetime import datetime, timezone
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from app import _load_or_update_input_set
from input_store import InputSet, _load_snapshot, _write_snapshot
from pricing import (
    CATALOG_PATH,
    PROPERTY_PATH,
    SCENARIO_FOLDERS,
    load_catalog,
    load_markers,
    load_properties,
)


class InputStoreParquetTests(unittest.TestCase):
    def test_input_set_round_trips_through_parquet(self):
        properties, duplicate_issues = load_properties(PROPERTY_PATH)
        catalog, _ = load_catalog(CATALOG_PATH)
        markers = {
            scenario: load_markers(scenario)
            for scenario in SCENARIO_FOLDERS
        }

        with tempfile.TemporaryDirectory() as temp_dir:
            snapshot = Path(temp_dir)
            _write_snapshot(
                snapshot,
                properties=properties,
                duplicate_issues=duplicate_issues,
                catalog=catalog,
                markers=markers,
            )
            (
                saved_properties,
                saved_issues,
                saved_catalog,
                saved_markers,
            ) = _load_snapshot(snapshot)

        self.assertEqual(saved_properties, properties)
        self.assertEqual(saved_issues, duplicate_issues)
        self.assertEqual(saved_catalog, catalog)
        self.assertEqual(saved_markers, markers)

    def test_calculation_reuses_the_current_parquet_set_without_reading_excel(self):
        properties, duplicate_issues = load_properties(PROPERTY_PATH)
        catalog, _ = load_catalog(CATALOG_PATH)
        markers = {
            scenario: load_markers(scenario)
            for scenario in SCENARIO_FOLDERS
        }
        current = InputSet(
            input_set_id="current",
            version=7,
            created_at=datetime.now(timezone.utc),
            properties=properties,
            duplicate_issues=duplicate_issues,
            catalog=catalog,
            markers=markers,
            properties_source="properties.parquet",
            catalog_source="catalog.parquet",
            marker_sources={
                scenario: marker_set.source_files
                for scenario, marker_set in markers.items()
            },
        )

        class EmptyForm:
            @staticmethod
            def get(_name):
                return None

        with (
            patch("app.get_current_input_set", return_value=current),
            patch("app.load_properties", side_effect=AssertionError("Excel reread")),
            patch("app.load_catalog", side_effect=AssertionError("Excel reread")),
            patch("app.load_markers", side_effect=AssertionError("Excel reread")),
        ):
            loaded, _ = asyncio.run(_load_or_update_input_set(EmptyForm()))

        self.assertIs(loaded, current)


if __name__ == "__main__":
    unittest.main()
