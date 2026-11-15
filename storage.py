from __future__ import annotations

from datetime import datetime
from io import BytesIO
import json
import os
from pathlib import Path
from threading import Lock
from typing import Mapping, Sequence

import pyarrow as pa
import pyarrow.parquet as pq
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from pricing import DATA_DIR


RESULTS_DIR = DATA_DIR.parent / "Resultados" / "Cartera"
_write_lock = Lock()


def write_calculation(
    rows: Sequence[Mapping[str, object]],
    *,
    run_id: str,
    scenario: str,
    calculated_at: str,
    calculation_mode: str = "validacion",
    input_set_version: int | None = None,
    source_files: Sequence[str],
    warnings: Sequence[str],
) -> Path:
    if not rows:
        raise ValueError("No hay resultados para guardar.")
    RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    output_path = RESULTS_DIR / f"{run_id}.parquet"
    temporary_path = RESULTS_DIR / f".{run_id}.tmp.parquet"
    result_rows = [
        {**row, "version_insumos": input_set_version}
        for row in rows
    ]
    table = pa.Table.from_pylist(result_rows)
    metadata = dict(table.schema.metadata or {})
    metadata.update(
        {
            b"title": b"Calculadora de precios de la GIOEE",
            b"logo": b"GIOEE",
            b"process": b"Cartera",
            b"scenario": scenario.encode("utf-8"),
            b"calculation_mode": calculation_mode.encode("utf-8"),
            b"input_set_version": str(input_set_version or "").encode("ascii"),
            b"calculated_at": calculated_at.encode("utf-8"),
        }
    )
    table = table.replace_schema_metadata(metadata)
    pq.write_table(table, temporary_path, compression="zstd")
    os.replace(temporary_path, output_path)

    log_path = RESULTS_DIR / "bitacora.parquet"
    log_temporary = RESULTS_DIR / ".bitacora.tmp.parquet"
    log_row = {
        "run_id": run_id,
        "fecha_calculo": calculated_at,
        "proceso": "Cartera",
        "escenario": scenario,
        "modo_calculo": calculation_mode,
        "version_insumos": input_set_version,
        "cantidad_registros": len(rows),
        "archivo_resultados": output_path.name,
        "archivos_insumo": json.dumps(list(source_files), ensure_ascii=False),
        "advertencias": json.dumps(list(warnings), ensure_ascii=False),
    }
    with _write_lock:
        new_log_table = pa.Table.from_pylist([log_row])
        if log_path.exists():
            old_log_table = pq.read_table(log_path)
            if "modo_calculo" not in old_log_table.column_names:
                old_log_table = old_log_table.append_column(
                    "modo_calculo",
                    pa.nulls(old_log_table.num_rows, type=pa.string()),
                )
            if "version_insumos" not in old_log_table.column_names:
                old_log_table = old_log_table.append_column(
                    "version_insumos",
                    pa.nulls(old_log_table.num_rows, type=pa.int64()),
                )
            new_log_table = new_log_table.select(old_log_table.column_names).cast(
                old_log_table.schema
            )
            new_log_table = pa.concat_tables([old_log_table, new_log_table])
        pq.write_table(new_log_table, log_temporary, compression="zstd")
        os.replace(log_temporary, log_path)
    return output_path


def calculation_file(run_id: str) -> Path | None:
    if not run_id or any(char not in "0123456789abcdef-" for char in run_id.lower()):
        return None
    path = RESULTS_DIR / f"{run_id}.parquet"
    return path if path.is_file() else None


def calculation_export_filename(
    calculated_at: str,
    scenario: str,
    input_set_version: int | None,
) -> str:
    try:
        calculation_date = datetime.fromisoformat(calculated_at).strftime("%Y%m%d")
    except ValueError:
        calculation_date = "fecha_desconocida"
    scenario_name = scenario.strip().casefold()
    version = f"{input_set_version:02d}" if input_set_version is not None else "sin_version"
    return f"precios_cartera_{scenario_name}_{calculation_date}_version_{version}.xlsx"


def calculation_xlsx(run_id: str) -> tuple[bytes, str] | None:
    path = calculation_file(run_id)
    if path is None:
        return None

    table = pq.read_table(path)
    rows = table.to_pylist()
    if not rows:
        return None

    fields = [
        ("Campo", "campo"),
        ("Hidrocarburo", "tipo_hidrocarburo"),
        ("Marcador", "tipo_gas"),
        ("API / BTU", "api_btu"),
        ("Unidad", "unidad"),
    ]
    year_columns = sorted(
        (name for name in table.column_names if name.startswith("AÑO ")),
        key=lambda name: int(name.removeprefix("AÑO ")),
    )
    fields.extend((name, name) for name in year_columns)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Precios Cartera"
    worksheet.append([label for label, _ in fields])
    for row in rows:
        worksheet.append([row.get(key) for _, key in fields])

    header_fill = PatternFill(fill_type="solid", fgColor="355D42")
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    worksheet.freeze_panes = "F2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for cell in worksheet["D"][1:]:
        cell.number_format = "0.00"
    for row in worksheet.iter_rows(min_row=2, min_col=6):
        for cell in row:
            cell.number_format = "0.0000"
    for column, width in enumerate((24, 18, 26, 14, 12), start=1):
        worksheet.column_dimensions[worksheet.cell(row=1, column=column).column_letter].width = width
    for column in range(6, len(fields) + 1):
        worksheet.column_dimensions[worksheet.cell(row=1, column=column).column_letter].width = 13

    metadata_sheet = workbook.create_sheet("Metadatos")
    first_row = rows[0]
    metadata = (
        ("Título", first_row.get("titulo_reporte")),
        ("Logo", first_row.get("logo")),
        ("Proceso", first_row.get("proceso")),
        ("Escenario", first_row.get("escenario")),
        ("Uso del cálculo", first_row.get("modo_calculo")),
        ("Versión de insumos", first_row.get("version_insumos")),
        ("Fecha de cálculo", first_row.get("fecha_calculo")),
    )
    metadata_sheet.append(["Dato", "Valor"])
    for item in metadata:
        metadata_sheet.append(item)
    for cell in metadata_sheet[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
    metadata_sheet.column_dimensions["A"].width = 24
    metadata_sheet.column_dimensions["B"].width = 48
    metadata_sheet.freeze_panes = "A2"

    output = BytesIO()
    workbook.save(output)
    filename = calculation_export_filename(
        str(first_row.get("fecha_calculo", "")),
        str(first_row.get("escenario", "")),
        first_row.get("version_insumos"),
    )
    return output.getvalue(), filename
