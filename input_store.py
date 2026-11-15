from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime
import json
import os
from pathlib import Path
import shutil
import uuid

import pyarrow as pa
import pyarrow.parquet as pq
import psycopg
from psycopg.types.json import Jsonb

from pricing import DATA_DIR, MarkerSet, Property


INPUTS_DIR = DATA_DIR.parent / "Resultados" / "Cartera" / "Insumos"


class InputStoreError(RuntimeError):
    """The PostgreSQL input registry or its Parquet snapshot is unavailable."""


@dataclass(frozen=True)
class InputSet:
    input_set_id: str
    version: int
    created_at: datetime
    properties: list[Property]
    duplicate_issues: list[str]
    catalog: set[str]
    markers: dict[str, MarkerSet]
    properties_source: str
    catalog_source: str
    marker_sources: dict[str, list[str]]


def _connect() -> psycopg.Connection:
    database_url = os.environ.get("DATABASE_URL", "").strip()
    if not database_url:
        raise InputStoreError(
            "Configura DATABASE_URL para conectar con PostgreSQL."
        )
    try:
        return psycopg.connect(database_url, connect_timeout=5)
    except psycopg.Error as exc:
        raise InputStoreError(
            "No se pudo conectar con PostgreSQL usando DATABASE_URL."
        ) from exc


def initialize_schema() -> None:
    try:
        with _connect() as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS input_sets (
                    input_set_id UUID PRIMARY KEY,
                    version INTEGER NOT NULL UNIQUE,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
                    properties_source TEXT NOT NULL,
                    catalog_source TEXT NOT NULL,
                    marker_sources JSONB NOT NULL,
                    duplicate_issues JSONB NOT NULL,
                    parquet_directory TEXT NOT NULL
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS current_input_set (
                    slot TEXT PRIMARY KEY CHECK (slot = 'cartera'),
                    input_set_id UUID NOT NULL REFERENCES input_sets(input_set_id),
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )
    except psycopg.Error as exc:
        raise InputStoreError("No se pudo inicializar el registro de insumos.") from exc


def _write_marker_set(path: Path, markers: MarkerSet) -> None:
    rows: list[dict[str, object]] = []
    for kind, series in (
        ("oil", markers.oil),
        ("gas_usd_per_mmbtu", markers.gas_usd_per_mmbtu),
        ("condensate_usd_per_bbl", markers.condensate_usd_per_bbl),
    ):
        for marker, yearly_values in series.items():
            rows.extend(
                {
                    "kind": kind,
                    "marker": marker,
                    "year": year,
                    "value": value,
                }
                for year, value in yearly_values.items()
            )
    rows.extend(
        {
            "kind": "gas_bn_usd_per_mpc",
            "marker": "Gas BN",
            "year": year,
            "value": value,
        }
        for year, value in markers.gas_bn_usd_per_mpc.items()
    )
    table = pa.Table.from_pylist(rows)
    table = table.replace_schema_metadata(
        {
            b"scenario": markers.scenario.encode("utf-8"),
            b"years": json.dumps(markers.years).encode("utf-8"),
            b"source_files": json.dumps(
                markers.source_files, ensure_ascii=False
            ).encode("utf-8"),
        }
    )
    pq.write_table(table, path, compression="zstd")


def _read_marker_set(path: Path) -> MarkerSet:
    table = pq.read_table(path)
    metadata = table.schema.metadata or {}
    scenario = metadata[b"scenario"].decode("utf-8")
    oil: dict[str, dict[int, float]] = {}
    gas: dict[str, dict[int, float]] = {}
    condensate: dict[str, dict[int, float]] = {}
    gas_bn: dict[int, float] = {}
    collections = {
        "oil": oil,
        "gas_usd_per_mmbtu": gas,
        "condensate_usd_per_bbl": condensate,
    }
    for row in table.to_pylist():
        kind = row["kind"]
        year = int(row["year"])
        value = float(row["value"])
        if kind == "gas_bn_usd_per_mpc":
            gas_bn[year] = value
        else:
            collections[kind].setdefault(row["marker"], {})[year] = value
    return MarkerSet(
        scenario=scenario,
        years=json.loads(metadata[b"years"].decode("utf-8")),
        oil=oil,
        gas_usd_per_mmbtu=gas,
        condensate_usd_per_bbl=condensate,
        gas_bn_usd_per_mpc=gas_bn,
        source_files=json.loads(metadata[b"source_files"].decode("utf-8")),
    )


def _write_snapshot(
    directory: Path,
    *,
    properties: list[Property],
    duplicate_issues: list[str],
    catalog: set[str],
    markers: dict[str, MarkerSet],
) -> None:
    properties_table = pa.Table.from_pylist([asdict(prop) for prop in properties])
    properties_table = properties_table.replace_schema_metadata(
        {
            b"duplicate_issues": json.dumps(
                duplicate_issues, ensure_ascii=False
            ).encode("utf-8")
        }
    )
    pq.write_table(properties_table, directory / "properties.parquet", compression="zstd")
    pq.write_table(
        pa.table({"field": sorted(catalog)}),
        directory / "catalog.parquet",
        compression="zstd",
    )
    for scenario, marker_set in markers.items():
        _write_marker_set(directory / f"markers_{scenario}.parquet", marker_set)


def _load_snapshot(directory: Path) -> tuple[
    list[Property], list[str], set[str], dict[str, MarkerSet]
]:
    properties_table = pq.read_table(directory / "properties.parquet")
    properties = [
        Property(**row) for row in properties_table.to_pylist()
    ]
    metadata = properties_table.schema.metadata or {}
    duplicate_issues = json.loads(
        metadata.get(b"duplicate_issues", b"[]").decode("utf-8")
    )
    catalog = set(pq.read_table(directory / "catalog.parquet")["field"].to_pylist())
    markers = {
        path.stem.removeprefix("markers_"): _read_marker_set(path)
        for path in directory.glob("markers_*.parquet")
    }
    return properties, duplicate_issues, catalog, markers


def create_input_set(
    *,
    properties: list[Property],
    duplicate_issues: list[str],
    catalog: set[str],
    markers: dict[str, MarkerSet],
    properties_source: str,
    catalog_source: str,
) -> InputSet:
    if not properties or not catalog or set(markers) != {"alto", "medio", "bajo"}:
        raise InputStoreError("El conjunto de insumos está incompleto.")

    input_set_id = uuid.uuid4()
    staging_directory = INPUTS_DIR / f".tmp-{input_set_id}"
    final_directory = INPUTS_DIR / str(input_set_id)
    INPUTS_DIR.mkdir(parents=True, exist_ok=True)
    staging_directory.mkdir()
    try:
        _write_snapshot(
            staging_directory,
            properties=properties,
            duplicate_issues=duplicate_issues,
            catalog=catalog,
            markers=markers,
        )
        os.replace(staging_directory, final_directory)
        marker_sources = {
            scenario: marker_set.source_files
            for scenario, marker_set in markers.items()
        }
        initialize_schema()
        with _connect() as connection:
            connection.execute("LOCK TABLE input_sets IN EXCLUSIVE MODE")
            version = connection.execute(
                "SELECT COALESCE(MAX(version), 0) + 1 FROM input_sets"
            ).fetchone()[0]
            row = connection.execute(
                """
                INSERT INTO input_sets (
                    input_set_id, version, properties_source, catalog_source,
                    marker_sources, duplicate_issues, parquet_directory
                ) VALUES (%s, %s, %s, %s, %s, %s, %s)
                RETURNING created_at
                """,
                (
                    input_set_id,
                    version,
                    properties_source,
                    catalog_source,
                    Jsonb(marker_sources),
                    Jsonb(duplicate_issues),
                    str(final_directory),
                ),
            ).fetchone()
            created_at = row[0]
            connection.execute(
                """
                INSERT INTO current_input_set (slot, input_set_id)
                VALUES ('cartera', %s)
                ON CONFLICT (slot) DO UPDATE
                SET input_set_id = EXCLUDED.input_set_id, updated_at = now()
                """,
                (input_set_id,),
            )
    except Exception as exc:
        shutil.rmtree(staging_directory, ignore_errors=True)
        shutil.rmtree(final_directory, ignore_errors=True)
        if isinstance(exc, (InputStoreError, OSError, ValueError)):
            raise
        if isinstance(exc, psycopg.Error):
            raise InputStoreError(
                "No se pudo registrar la nueva versión de insumos en PostgreSQL."
            ) from exc
        raise

    return InputSet(
        input_set_id=str(input_set_id),
        version=version,
        created_at=created_at,
        properties=properties,
        duplicate_issues=duplicate_issues,
        catalog=catalog,
        markers=markers,
        properties_source=properties_source,
        catalog_source=catalog_source,
        marker_sources=marker_sources,
    )


def get_current_input_set() -> InputSet | None:
    initialize_schema()
    try:
        with _connect() as connection:
            row = connection.execute(
                """
                SELECT s.input_set_id, s.version, s.created_at,
                       s.properties_source, s.catalog_source,
                       s.marker_sources, s.parquet_directory
                FROM current_input_set c
                JOIN input_sets s ON s.input_set_id = c.input_set_id
                WHERE c.slot = 'cartera'
                """
            ).fetchone()
    except psycopg.Error as exc:
        raise InputStoreError("No se pudo consultar la versión vigente de insumos.") from exc
    if row is None:
        return None

    input_set_id, version, created_at, properties_source, catalog_source, marker_sources, directory = row
    snapshot_directory = Path(directory)
    try:
        properties, duplicate_issues, catalog, markers = _load_snapshot(
            snapshot_directory
        )
    except (OSError, pa.ArrowException, KeyError, ValueError) as exc:
        raise InputStoreError(
            f"No se pudo leer el set Parquet vigente v{version}."
        ) from exc
    return InputSet(
        input_set_id=str(input_set_id),
        version=version,
        created_at=created_at,
        properties=properties,
        duplicate_issues=duplicate_issues,
        catalog=catalog,
        markers=markers,
        properties_source=properties_source,
        catalog_source=catalog_source,
        marker_sources=marker_sources,
    )
