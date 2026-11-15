from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
import math
import re
import unicodedata
from pathlib import Path
from typing import BinaryIO, Iterable, Sequence

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "Data_Informacion"
PROPERTY_PATH = (
    DATA_DIR
    / "Propiedades"
    / "PropiedadesCartera-2025"
    / "propiedadesCartera_2025.xlsx"
)
CATALOG_PATH = DATA_DIR / "Catalogo" / "06_Campo.xlsx"
MARKERS_DIR = DATA_DIR / "Marcadores" / "MaracdoresCartera-2025"

REPORT_TITLE = "Calculadora de precios de la GIOEE"
REPORT_LOGO = "GIOEE"
PROCESS_NAME = "Cartera"
REPORT_START_YEAR = 2025
REPORT_END_YEAR = 2124
MARKER_END_YEAR = 2050
OIL_API_CAP = 48.5
CALCULATION_MODES = {
    "validacion": "Validación de precios por campo",
    "reporte_final": "Reporte final de proceso",
}
HEAVY_CRUDE_RELATIONS = {
    "Kayab": (8.6, 0.4238),
    "Ayatsil": (10.2, 0.3284),
    "Utsil": (10.2, 0.3284),
    "Pit": (10.8, 0.2932),
    "Tekel": (10.9, 0.2874),
}
OIL_REFERENCE_FIELDS = {
    "Homol": 37.1,
    "Ket": 15.37,
    "Kayab": 8.6,
}
SCENARIO_FOLDERS = {
    "alto": "marcadores_250225_alto",
    "bajo": "marcadores_250225_bajo",
    "medio": "marcadores_250225_medio",
}
GAS_MARKER_TYPES = (
    "GH Dulce Norte",
    "GH Amargo Golfo",
    "GH Amargo Sureste",
    "GH Dulce Sureste",
)
PROPERTY_HEADERS = (
    "Campo",
    "API_Poder_Calorífico",
    "Hidrocarburo",
    "Unidad",
    "Tipo_gas_condensado",
    "Calculado_como_Aceite",
)


class InputError(ValueError):
    """An input file is unreadable or does not meet the supported schema."""


@dataclass(frozen=True)
class Property:
    field: str
    api_btu: float
    hydrocarbon: str
    unit: str
    marker_type: str
    calculated_as_oil: float | None = None


@dataclass
class MarkerSet:
    scenario: str
    years: list[int]
    oil: dict[str, dict[int, float]]
    gas_usd_per_mmbtu: dict[str, dict[int, float]]
    condensate_usd_per_bbl: dict[str, dict[int, float]]
    gas_bn_usd_per_mpc: dict[int, float]
    source_files: list[str]


def normalize_text(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return " ".join(text.casefold().strip().split())


def normalize_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]", "", normalize_text(value))


def normalize_field_name(value: object) -> str:
    return normalize_text(value).replace("-", " ")


def _read_sheet_rows(source: Path | bytes, sheet_index: int = 0) -> list[tuple]:
    try:
        workbook = load_workbook(
            BytesIO(source) if isinstance(source, bytes) else source,
            read_only=True,
            data_only=True,
        )
    except (OSError, ValueError) as exc:
        raise InputError(f"No se pudo abrir el archivo Excel: {exc}") from exc
    try:
        if sheet_index >= len(workbook.worksheets):
            raise InputError("El archivo no contiene la hoja requerida.")
        return list(workbook.worksheets[sheet_index].iter_rows(values_only=True))
    finally:
        workbook.close()


def _header_positions(headers: Sequence[object]) -> dict[str, int]:
    return {
        normalize_header(header): index
        for index, header in enumerate(headers)
        if header is not None
    }


def _as_number(value: object, description: str) -> float:
    if isinstance(value, bool) or value is None:
        raise InputError(f"El valor de {description} debe ser numérico.")
    if isinstance(value, (int, float)):
        number = float(value)
    else:
        text = str(value).strip()
        if not text:
            raise InputError(f"El valor de {description} está vacío.")
        if "," in text and "." not in text:
            text = text.replace(",", ".")
        try:
            number = float(text)
        except ValueError as exc:
            raise InputError(f"El valor de {description} no es numérico.") from exc
    if not math.isfinite(number):
        raise InputError(f"El valor de {description} no es finito.")
    return number


def load_properties(source: Path | bytes) -> tuple[list[Property], list[str]]:
    rows = _read_sheet_rows(source)
    if not rows:
        raise InputError("El archivo de propiedades está vacío.")
    positions = _header_positions(rows[0])
    required = {
        normalize_header(name): name for name in PROPERTY_HEADERS[:-1]
    }
    missing = [label for key, label in required.items() if key not in positions]
    oil_flag_key = normalize_header(PROPERTY_HEADERS[-1])
    if oil_flag_key not in positions:
        missing.append(PROPERTY_HEADERS[-1])
    if missing:
        raise InputError(
            "Faltan columnas requeridas en Propiedades: " + ", ".join(missing)
        )

    properties: list[Property] = []
    issues: list[str] = []
    seen: set[tuple[str, str]] = set()
    for row_number, row in enumerate(rows[1:], start=2):
        if not any(value is not None for value in row):
            continue
        field = str(row[positions["campo"]] or "").strip()
        hydrocarbon = str(row[positions["hidrocarburo"]] or "").strip()
        marker_type = str(row[positions["tipogascondensado"]] or "").strip()
        unit = str(row[positions["unidad"]] or "").strip().lower()
        if not field or not hydrocarbon or not marker_type or not unit:
            raise InputError(
                f"Hay un dato obligatorio vacío en la fila {row_number} de Propiedades."
            )
        value = _as_number(
            row[positions["apipodercalorifico"]],
            f"API/Poder Calorífico en la fila {row_number}",
        )
        if value < 0:
            raise InputError(
                f"API/Poder Calorífico no puede ser negativo (fila {row_number})."
            )
        flag_value = row[positions[oil_flag_key]]
        calculated_as_oil = (
            None
            if flag_value is None or str(flag_value).strip() == ""
            else _as_number(flag_value, f"Calculado_como_Aceite en fila {row_number}")
        )
        hydro_key = normalize_text(hydrocarbon)
        expected_units = {
            "aceite": "usd/b",
            "condensado": "usd/b",
            "gas": "usd/mpc",
        }
        if hydro_key not in expected_units:
            raise InputError(
                f"Hidrocarburo no reconocido en la fila {row_number}: {hydrocarbon}."
            )
        if unit != expected_units[hydro_key]:
            raise InputError(
                f"Unidad inesperada en la fila {row_number}: {unit}; "
                f"se esperaba {expected_units[hydro_key]}."
            )
        key = (normalize_field_name(field), hydro_key)
        if key in seen:
            issues.append(
                f"Propiedad duplicada para {field} / {hydrocarbon}; "
                "se conservará la primera fila, como en el modelo 2025."
            )
        else:
            seen.add(key)
            properties.append(
                Property(
                    field=field,
                    api_btu=value,
                    hydrocarbon=hydrocarbon,
                    unit=unit,
                    marker_type=marker_type,
                    calculated_as_oil=calculated_as_oil,
                )
            )
    if not properties:
        raise InputError("No se encontraron registros de propiedades.")
    return properties, issues


def load_catalog(source: Path | bytes) -> tuple[set[str], list[str]]:
    rows = _read_sheet_rows(source)
    if not rows:
        raise InputError("El catálogo está vacío.")
    positions = _header_positions(rows[0])
    field_key = normalize_header("Campo")
    if field_key not in positions:
        raise InputError("El catálogo debe tener una columna 'Campo'.")
    catalog: set[str] = set()
    display_names: dict[str, str] = {}
    for row in rows[1:]:
        if positions[field_key] >= len(row):
            continue
        value = str(row[positions[field_key]] or "").strip()
        if value:
            normalized = normalize_field_name(value)
            if normalized in catalog:
                continue
            catalog.add(normalized)
            display_names[normalized] = value
    if not catalog:
        raise InputError("El catálogo no contiene nombres de campo.")
    return catalog, list(display_names.values())


def _find_row(
    rows: Sequence[Sequence[object]],
    *,
    code: str | None = None,
    label: str | None = None,
) -> Sequence[object]:
    expected_code = normalize_text(code) if code else None
    expected_label = normalize_text(label) if label else None
    for row in rows:
        if expected_code and any(normalize_text(cell) == expected_code for cell in row):
            return row
        if expected_label and any(
            normalize_text(cell) == expected_label for cell in row
        ):
            return row
    criterion = code or label or "dato"
    raise InputError(f"No se encontró el marcador requerido: {criterion}.")


def _row_has_label(row: Sequence[object], label: str) -> bool:
    expected = normalize_text(label)
    return any(normalize_text(cell) == expected for cell in row)


def _price_by_year(
    row: Sequence[object],
    year_positions: dict[int, int],
    description: str,
    *,
    factor: float = 1.0,
) -> dict[int, float]:
    if factor <= 0 or not math.isfinite(factor):
        raise InputError(f"El factor de conversión no es válido para {description}.")
    prices: dict[int, float] = {}
    for year, column in year_positions.items():
        if column >= len(row):
            raise InputError(f"Falta {description} para {year}.")
        prices[year] = _as_number(row[column], f"{description}, año {year}") / factor
    return prices


def _year_positions(rows: Sequence[Sequence[object]], row_number: int) -> dict[int, int]:
    row_index = row_number - 1
    if row_index >= len(rows):
        raise InputError("No se encontró el encabezado anual de los marcadores.")
    result = {
        int(value): index
        for index, value in enumerate(rows[row_index])
        if isinstance(value, (int, float))
        and not isinstance(value, bool)
        and int(value) == value
        and REPORT_START_YEAR <= int(value) <= REPORT_END_YEAR
    }
    if not result:
        raise InputError("El archivo de marcadores no contiene años reconocidos.")
    return result


def _scenario_file(folder: Path, prefix: str) -> Path:
    candidates = [
        path
        for path in folder.glob("*.xlsx")
        if path.name[:2].upper() == prefix and not path.name.startswith("~$")
    ]
    if len(candidates) != 1:
        raise InputError(
            f"El escenario debe contener exactamente un archivo {prefix}. "
            f"Se encontraron {len(candidates)}."
        )
    return candidates[0]


def load_markers(scenario: str) -> MarkerSet:
    scenario_key = normalize_text(scenario)
    if scenario_key not in SCENARIO_FOLDERS:
        raise InputError("Selecciona un escenario Alto, Medio o Bajo.")
    folder = MARKERS_DIR / SCENARIO_FOLDERS[scenario_key]
    if not folder.is_dir():
        raise InputError(f"No existe la carpeta de marcadores: {folder.name}.")

    paths = {prefix: _scenario_file(folder, prefix) for prefix in ("VE", "VS", "VG")}
    ve_rows = _read_sheet_rows(paths["VE"])
    vs_rows = _read_sheet_rows(paths["VS"])
    vg_rows = _read_sheet_rows(paths["VG"])
    ve_years = _year_positions(ve_rows, 4)
    vs_years = _year_positions(vs_rows, 5)
    vg_years = _year_positions(vg_rows, 4)
    years = sorted(set(ve_years) & set(vs_years) & set(vg_years))
    if not years or years[0] != REPORT_START_YEAR:
        raise InputError(
            f"Los archivos VE, VS y VG deben compartir años desde {REPORT_START_YEAR}."
        )

    ve_years = {year: ve_years[year] for year in years}
    vs_years = {year: vs_years[year] for year in years}
    vg_years = {year: vg_years[year] for year in years}

    oil: dict[str, dict[int, float]] = {}
    for marker, label in (
        ("Istmo", "Istmo"),
        ("Maya", "Maya"),
        ("Olmeca", "Olmeca"),
    ):
        row = _find_row(vs_rows, label=label)
        oil[marker] = _price_by_year(row, vs_years, f"Marcador VS {marker}")
    tci_row = _find_row(ve_rows, code="CRUTCICTCI")
    tci = _price_by_year(tci_row, ve_years, "Tipo de cambio TCI")
    altamira_row = next(
        (
            row
            for row in ve_rows
            if len(row) > 3
            and normalize_text(row[0]) == ""
            and normalize_text(row[1]) == ""
            and normalize_text(row[2]) == "us$/b"
            and normalize_text(row[3]) == "crudo altamira"
        ),
        None,
    )
    if altamira_row is None:
        raise InputError("No se encontró el marcador de Altamira de Ventas a PMI.")
    altamira_pesos = _price_by_year(
        altamira_row, ve_years, "Marcador VE Altamira de Ventas a PMI"
    )
    oil["Altamira"] = {}
    for year in years:
        if tci[year] <= 0:
            raise InputError(f"El TCI no es positivo para {year}.")
        oil["Altamira"][year] = altamira_pesos[year] / tci[year]

    sour_condensate_row = _find_row(ve_rows, code="CRUPGPC910122")
    sweet_condensate_row = _find_row(ve_rows, code="CRUPGPC910124")
    condensate = {
        "Condensado Amargo": _price_by_year(
            sour_condensate_row, ve_years, "Condensado amargo", factor=1.0
        ),
        "Condensado Dulce": _price_by_year(
            sweet_condensate_row, ve_years, "Condensado dulce", factor=1.0
        ),
    }
    for name in condensate:
        for year in years:
            if tci[year] <= 0:
                raise InputError(f"El TCI no es positivo para {year}.")
            condensate[name][year] /= tci[year]

    gas_usd_per_mmbtu: dict[str, dict[int, float]] = {}
    gas_labels = {
        "GH Dulce Norte": "GH Dulce Norte, US$/Mpc",
        "GH Amargo Golfo": "GH Amargo Golfo, US$/Mpc",
        "GH Amargo Sureste": "GH Amargo Sureste, US$/Mpc",
        "GH Dulce Sureste": "GH Dulce Sureste, US$/Mpc",
    }
    for marker, label in gas_labels.items():
        row = next(
            (candidate for candidate in ve_rows if len(candidate) > 3 and _row_has_label(candidate, label)),
            None,
        )
        if row is None:
            raise InputError(f"No se encontró el marcador de gas {marker}.")
        factor = _as_number(row[0], f"Factor MMBTU/MPC de {marker}")
        gas_usd_per_mmbtu[marker] = _price_by_year(
            row, ve_years, marker, factor=factor
        )

    gas_bn_row = _find_row(vg_rows, code="GLPEPC930109")
    tcp_row = _find_row(vg_rows, code="GLTCCTCP")
    gas_bn_pesos = _price_by_year(gas_bn_row, vg_years, "Gas BN PEP")
    tcp = _price_by_year(tcp_row, vg_years, "Tipo de cambio TCP")
    gas_bn: dict[int, float] = {}
    for year in years:
        if tcp[year] <= 0:
            raise InputError(f"El TCP no es positivo para {year}.")
        gas_bn[year] = gas_bn_pesos[year] / tcp[year]

    if max(years) < MARKER_END_YEAR:
        raise InputError(
            f"Los archivos de referencia deben incluir el año {MARKER_END_YEAR}."
        )
    return MarkerSet(
        scenario=scenario_key,
        years=years,
        oil=oil,
        gas_usd_per_mmbtu=gas_usd_per_mmbtu,
        condensate_usd_per_bbl=condensate,
        gas_bn_usd_per_mpc=gas_bn,
        source_files=[path.name for path in paths.values()],
    )


def _market_value(values: dict[int, float], year: int) -> float:
    if year in values:
        return values[year]
    available_years = sorted(values)
    if year > available_years[-1]:
        return values[available_years[-1]]
    raise InputError(f"Falta un marcador anual para {year}.")


def _linear_price(
    api: float,
    lower_api: float,
    lower_price: float,
    upper_api: float,
    upper_price: float,
) -> float:
    if upper_api == lower_api:
        raise InputError("No se puede interpolar entre API idénticos.")
    return lower_price + (api - lower_api) * (
        (upper_price - lower_price) / (upper_api - lower_api)
    )


def oil_price(api: float, markers: MarkerSet, year: int) -> float:
    if api == 0:
        return 0.0
    api = min(api, OIL_API_CAP)
    istmo = _market_value(markers.oil["Istmo"], year)
    maya = _market_value(markers.oil["Maya"], year)
    olmeca = _market_value(markers.oil["Olmeca"], year)
    altamira = _market_value(markers.oil["Altamira"], year)

    homol_price = _linear_price(37.1, 33.6, istmo, 39.3, olmeca)
    blend_price = _linear_price(15.37, 15.5, altamira, 22.0, maya)
    anchors = [
        (15.37, blend_price),
        (15.5, altamira),
        (22.0, maya),
        (33.6, istmo),
        (37.1, homol_price),
        (39.3, olmeca),
    ]
    for field, (api_value, ratio) in HEAVY_CRUDE_RELATIONS.items():
        heavy_price = blend_price + ratio * (blend_price - homol_price)
        anchors.append((api_value, heavy_price))
    anchors.sort(key=lambda anchor: anchor[0])

    for (lower_api, lower_price), (upper_api, upper_price) in zip(
        anchors, anchors[1:]
    ):
        if lower_api <= api <= upper_api:
            return _linear_price(api, lower_api, lower_price, upper_api, upper_price)
    if api < anchors[0][0]:
        if api < 8.6:
            raise InputError("El API está por debajo del límite Kayab (8.6).")
        return _linear_price(api, *anchors[0], *anchors[1])
    return _linear_price(api, *anchors[-2], *anchors[-1])


def _property_by_type(
    properties: Iterable[Property],
) -> dict[str, dict[str, Property]]:
    grouped: dict[str, dict[str, Property]] = {}
    for prop in properties:
        field_key = normalize_field_name(prop.field)
        hydro_key = normalize_text(prop.hydrocarbon)
        grouped.setdefault(field_key, {})
        grouped[field_key].setdefault(hydro_key, prop)
    return grouped


def validate_properties(
    properties: Sequence[Property],
    catalog: set[str],
    duplicate_issues: Sequence[str] = (),
) -> tuple[list[str], list[str]]:
    errors: list[str] = []
    warnings = list(duplicate_issues)
    grouped = _property_by_type(properties)
    expected_hydrocarbons = {"aceite", "condensado", "gas"}
    non_catalog = sorted(
        {prop.field for prop in properties if normalize_field_name(prop.field) not in catalog}
    )
    if non_catalog:
        preview = ", ".join(non_catalog[:12])
        more = f" y {len(non_catalog) - 12} más" if len(non_catalog) > 12 else ""
        warnings.append(
            f"{len(non_catalog)} campos no aparecen en el catálogo y no se incluirán "
            f"en los resultados: {preview}{more}."
        )

    zero_fields: list[str] = []
    for field_key, products in grouped.items():
        missing = expected_hydrocarbons - products.keys()
        if missing:
            field_name = next(prop.field for prop in properties if normalize_field_name(prop.field) == field_key)
            errors.append(
                f"{field_name} no tiene los tres registros de hidrocarburos: "
                + ", ".join(sorted(missing))
                + "."
            )
            continue
        field_props = list(products.values())
        if not any(prop.api_btu != 0 for prop in field_props):
            zero_fields.append(field_props[0].field)
        oil_prop = products["aceite"]
        condensate_prop = products["condensado"]
        if (
            oil_prop.api_btu
            and oil_prop.api_btu < 8.6
        ) or (
            condensate_prop.calculated_as_oil == 13
            and condensate_prop.api_btu
            and condensate_prop.api_btu < 8.6
        ):
            errors.append(
                f"{field_props[0].field} tiene un API inferior al límite Kayab (8.6)."
            )
    if zero_fields:
        warnings.append(
            f"{len(zero_fields)} campos tienen API/Poder Calorífico igual a cero "
            "en los tres hidrocarburos; se omitirán, como en los entregables 2025."
        )
    if not grouped:
        errors.append("No se encontraron campos para calcular.")
    return errors, warnings


def calculate_prices(
    properties: Sequence[Property],
    catalog: set[str],
    markers: MarkerSet,
    *,
    calculated_at: str,
    run_id: str,
    calculation_mode: str = "validacion",
) -> list[dict[str, object]]:
    if calculation_mode not in CALCULATION_MODES:
        raise InputError("Selecciona Validación o Reporte final de proceso.")
    grouped = _property_by_type(properties)
    catalog_names = set(catalog)
    years = range(REPORT_START_YEAR, REPORT_END_YEAR + 1)
    results: list[dict[str, object]] = []
    for field_key, products in grouped.items():
        if field_key not in catalog_names:
            continue
        if not {"aceite", "condensado", "gas"}.issubset(products):
            continue
        oil_prop = products["aceite"]
        condensate_prop = products["condensado"]
        gas_prop = products["gas"]
        if not any(prop.api_btu != 0 for prop in products.values()):
            continue

        report_products = (
            (oil_prop, "Aceite"),
            (condensate_prop, "Condensado"),
            (gas_prop, "Gas"),
        )
        for prop, product_type in report_products:
            prices: dict[int, float] = {}
            for year in years:
                if product_type == "Aceite":
                    price = oil_price(prop.api_btu, markers, year)
                elif (
                    product_type == "Condensado"
                    and prop.calculated_as_oil == 13
                    and not (
                        calculation_mode == "reporte_final"
                        and field_key == normalize_field_name("Nejo")
                    )
                ):
                    price = oil_price(prop.api_btu, markers, year)
                elif product_type == "Condensado":
                    if prop.api_btu == 0:
                        price = 0.0
                    else:
                        try:
                            price = _market_value(
                                markers.condensate_usd_per_bbl[prop.marker_type],
                                year,
                            )
                        except KeyError as exc:
                            raise InputError(
                                f"Marcador de condensado no reconocido: {prop.marker_type}."
                            ) from exc
                else:
                    if prop.api_btu == 0:
                        price = 0.0
                    else:
                        try:
                            marker = _market_value(
                                markers.gas_usd_per_mmbtu[prop.marker_type],
                                year,
                            )
                        except KeyError as exc:
                            raise InputError(
                                f"Marcador de gas no reconocido: {prop.marker_type}."
                            ) from exc
                        price = prop.api_btu * marker / 1000
                prices[year] = price
            results.append(
                _report_row(
                    prop,
                    product_type,
                    prop.marker_type,
                    prices,
                    markers.scenario,
                    calculated_at,
                    run_id,
                    calculation_mode,
                )
            )

        gas_bn_prices = {
            year: _market_value(markers.gas_bn_usd_per_mpc, year) for year in years
        }
        results.append(
            _report_row(
                gas_prop,
                "Gas BN",
                "Gas BN",
                gas_bn_prices,
                markers.scenario,
                calculated_at,
                run_id,
                calculation_mode,
            )
        )
    return results


def _report_row(
    prop: Property,
    product_type: str,
    marker_type: str,
    prices: dict[int, float],
    scenario: str,
    calculated_at: str,
    run_id: str,
    calculation_mode: str,
) -> dict[str, object]:
    row: dict[str, object] = {
        "run_id": run_id,
        "fecha_periodo": REPORT_START_YEAR,
        "escenario": scenario.title(),
        "modo_calculo": calculation_mode,
        "campo": prop.field,
        "tipo_gas": marker_type,
        "api_btu": prop.api_btu,
        "tipo_hidrocarburo": product_type,
        "unidad": prop.unit,
        "titulo_reporte": REPORT_TITLE,
        "logo": REPORT_LOGO,
        "proceso": PROCESS_NAME,
        "fecha_calculo": calculated_at,
    }
    row.update(
        {
            f"AÑO {year}": _market_value(prices, year)
            for year in range(REPORT_START_YEAR, REPORT_END_YEAR + 1)
        }
    )
    return row


def property_distribution(
    properties: Sequence[Property], catalog: set[str]
) -> dict[str, list[tuple[str, float]]]:
    grouped = _property_by_type(properties)
    result: dict[str, list[tuple[str, float]]] = {
        "Aceite (°API)": [],
        "Condensado (°API)": [],
        "Gas (BTU/PC)": [],
    }
    for products in grouped.values():
        field = next(iter(products.values())).field
        if normalize_field_name(field) not in catalog:
            continue
        if not {"aceite", "condensado", "gas"}.issubset(products):
            continue
        result["Aceite (°API)"].append((field, products["aceite"].api_btu))
        result["Condensado (°API)"].append(
            (field, products["condensado"].api_btu)
        )
        result["Gas (BTU/PC)"].append((field, products["gas"].api_btu))
    return result
