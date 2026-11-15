from __future__ import annotations

from datetime import datetime
import html
from pathlib import Path
import uuid

from fasthtml.common import (
    Aside,
    Button,
    Div,
    Details,
    Form,
    H1,
    H2,
    H3,
    Header,
    Input,
    Label,
    Link,
    Main,
    Option,
    P,
    Script,
    Section,
    Select,
    Small,
    Span,
    Strong,
    Summary,
    Table,
    Tbody,
    Td,
    Textarea,
    Th,
    Thead,
    Title,
    Tr,
    Titled,
    Ul,
    Li,
    fast_app,
)
from starlette.requests import Request
from starlette.responses import Response
from starlette.staticfiles import StaticFiles
from starlette.datastructures import UploadFile

from charts import distribution_chart, marker_chart, oil_marker_series
from input_store import (
    InputSet,
    InputStoreError,
    create_input_set,
    get_current_input_set,
)
from pricing import (
    CATALOG_PATH,
    CALCULATION_MODES,
    GAS_MARKER_TYPES,
    MARKERS_DIR,
    PROPERTY_PATH,
    REPORT_END_YEAR,
    REPORT_START_YEAR,
    REPORT_TITLE,
    SCENARIO_FOLDERS,
    InputError,
    MarkerSet,
    calculate_prices,
    load_catalog,
    load_markers,
    load_properties,
    property_distribution,
    validate_properties,
)
from storage import (
    calculation_export_filename,
    calculation_xlsx,
    write_calculation,
)


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"

app, rt = fast_app(
    hdrs=(
        Link(rel="stylesheet", href="/static/app.css?v=12"),
        Script(src="/static/app.js?v=2", defer=True),
    )
)
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


def _logo() -> object:
    return Div(
        Span("G", cls="logo-mark", aria_hidden="true"),
        Div(
            Strong("GIOEE"),
            Small("Precios por campo"),
            cls="logo-copy",
        ),
        cls="brand",
        role="img",
        aria_label="GIOEE - Precios por campo",
    )


def _sidebar() -> object:
    return Details(
        Summary("Tipos de cálculo de precios"),
        Div(
            Div(
                "Precios Cartera",
                cls="process-panel active",
                aria_current="page",
            ),
            Details(
                Summary("Precios para PDEs"),
                P("Pendiente de definir e implementar sus insumos y reglas."),
                cls="process-panel pending",
            ),
            Details(
                Summary("Flujos de Deterioro"),
                P("Pendiente de definir e implementar sus insumos y reglas."),
                cls="process-panel pending",
            ),
            Details(
                Summary("Cartera Exploración"),
                P("Pendiente de definir e implementar sus insumos y reglas."),
                cls="process-panel pending",
            ),
            Details(
                Summary("Programa Operativo"),
                P("Pendiente de definir e implementar sus insumos y reglas."),
                cls="process-panel pending",
            ),
            cls="process-list",
        ),
        open=True,
        cls="process-navigation",
    )


def _upload_field(label: str, name: str, default_path: Path) -> object:
    return Div(
        Label(label, fr=name),
        Input(
            type="file",
            name=name,
            id=name,
            accept=".xlsx",
        ),
        Small(
            "Sin archivo nuevo se conserva el set vigente; el set inicial usa "
            f"{default_path.relative_to(ROOT)}."
        ),
        cls="file-field",
    )


async def _read_optional_upload(
    form: object,
    field_name: str,
) -> tuple[bytes | None, str | None]:
    upload = form.get(field_name)
    if not isinstance(upload, UploadFile) or not upload.filename:
        return None, None
    if Path(upload.filename).suffix.lower() != ".xlsx":
        raise InputError(f"{upload.filename}: solo se admiten libros .xlsx.")
    data = await upload.read()
    if not data:
        raise InputError(f"El archivo cargado {upload.filename} está vacío.")
    return data, Path(upload.filename).name


async def _load_or_update_input_set(form: object) -> tuple[InputSet, list[str]]:
    current = get_current_input_set()
    properties_data, properties_filename = await _read_optional_upload(
        form, "properties_file"
    )
    catalog_data, catalog_filename = await _read_optional_upload(form, "catalog_file")

    if current is None:
        properties, duplicate_issues = load_properties(
            properties_data if properties_data is not None else PROPERTY_PATH
        )
        catalog, _ = load_catalog(
            catalog_data if catalog_data is not None else CATALOG_PATH
        )
        markers = {
            scenario: load_markers(scenario)
            for scenario in SCENARIO_FOLDERS
        }
        properties_source = properties_filename or str(PROPERTY_PATH.relative_to(ROOT))
        catalog_source = catalog_filename or str(CATALOG_PATH.relative_to(ROOT))
    else:
        properties = current.properties
        duplicate_issues = current.duplicate_issues
        catalog = current.catalog
        markers = current.markers
        properties_source = current.properties_source
        catalog_source = current.catalog_source
        if properties_data is not None:
            properties, duplicate_issues = load_properties(properties_data)
            properties_source = properties_filename or properties_source
        if catalog_data is not None:
            catalog, _ = load_catalog(catalog_data)
            catalog_source = catalog_filename or catalog_source

    errors, warnings = validate_properties(properties, catalog, duplicate_issues)
    if errors:
        raise InputError(" ".join(errors))

    if (
        current is None
        or properties_data is not None
        or catalog_data is not None
    ):
        current = create_input_set(
            properties=properties,
            duplicate_issues=duplicate_issues,
            catalog=catalog,
            markers=markers,
            properties_source=properties_source,
            catalog_source=catalog_source,
        )
    return current, warnings


def _input_form() -> object:
    scenario_options = [
        Option("Selecciona un escenario", value="", selected=True, disabled=True),
        Option("Alto", value="alto"),
        Option("Medio", value="medio"),
        Option("Bajo", value="bajo"),
    ]
    calculation_mode_options = [
        Option(
            label,
            value=value,
            selected=value == "reporte_final",
        )
        for value, label in CALCULATION_MODES.items()
    ]
    return Form(
        Div(
            Label("Escenario de marcadores", fr="scenario"),
            Select(*scenario_options, name="scenario", id="scenario", required=True),
            Small(
                "Se cargan juntos VE, VS y VG desde la carpeta del escenario seleccionado."
            ),
            cls="file-field",
        ),
        Div(
            Label("Uso del cálculo", fr="calculation_mode"),
            Select(
                *calculation_mode_options,
                name="calculation_mode",
                id="calculation_mode",
                required=True,
            ),
            Small(
                "En el reporte final, Nejo usa el marcador de condensado directamente; "
                "el gas conserva su conversión por poder calorífico."
            ),
            cls="file-field",
        ),
        _upload_field("Propiedades de hidrocarburos (opcional)", "properties_file", PROPERTY_PATH),
        _upload_field("Catálogo de campos (opcional)", "catalog_file", CATALOG_PATH),
        Button("Validar insumos y calcular", type="submit", cls="primary-button"),
        action="/calculate",
        method="post",
        enctype="multipart/form-data",
        hx_post="/calculate",
        hx_target="#calculation-result",
        hx_swap="innerHTML",
        cls="input-form",
    )


def _page(content: object | None = None, *, input_status: str | None = None) -> object:
    page_body = Div(
        Header(_logo(), Div("PostgreSQL local", cls="header-note"), cls="topbar"),
        Div(
            Aside(_sidebar(), cls="sidebar"),
            Main(
                Section(
                    Div(
                        H1("Calculadora de precios de la GIOEE"),
                        P(
                            "Estimación anual de aceite, gas y condensado por campo. "
                            "La primera etapa replica el proceso de Cartera 2025.",
                            cls="intro",
                        ),
                        cls="page-heading",
                    ),
                    Section(
                        H2("Tipo de cálculo: Precios Cartera"),
                        P(
                            "Insumos exclusivos para Cartera. Los marcadores VE, VS y VG "
                            "se cargan desde el escenario seleccionado.",
                            cls="section-intro",
                        ),
                        P(
                            "El cálculo de Cartera usa el ModeloCartera-2025 como "
                            "referencia y no modifica los archivos fuente.",
                            cls="muted",
                        ),
                        P(input_status, cls="input-set-status") if input_status else None,
                        Details(
                            Summary("Opciones de carga y cálculo"),
                            _input_form(),
                            open=True,
                            id="calculation-options",
                            cls="input-options",
                        ),
                        cls="panel calculation-panel",
                    ),
                    Section(
                        Div(content or P("Selecciona escenario e insumos para revisar y calcular.", cls="muted"), id="calculation-result"),
                        cls="result-area",
                    ),
                    cls="main-content",
                ),
                cls="workspace",
            ),
            cls="app-layout",
        ),
        cls="app-shell",
    )
    return Titled(REPORT_TITLE, page_body)


def _issue_list(title: str, items: list[str], style: str) -> object:
    if not items:
        return Div()
    return Details(
        Summary(f"{title} ({len(items)})"),
        Ul(*(Li(item) for item in items)),
        cls=f"issue-list {style}",
    )


def _chart_panel(
    markers: MarkerSet,
    distributions: dict[str, list[tuple[str, float]]],
) -> object:
    oil_series = oil_marker_series(markers)
    gas_series = {
        name: [values[year] for year in markers.years]
        for name, values in markers.gas_usd_per_mmbtu.items()
    }
    condensate_series = {
        name: [values[year] for year in markers.years]
        for name, values in markers.condensate_usd_per_bbl.items()
    }
    gas_bn_series = {
        "Gas BN": [markers.gas_bn_usd_per_mpc[year] for year in markers.years]
    }
    charts = [
        ("Aceite", marker_chart("Marcadores de aceite y referencias", "USD/b", markers.years, oil_series)),
        ("Gas", marker_chart("Marcadores de gas", "USD/MMBTU", markers.years, gas_series)),
        ("Condensado", marker_chart("Marcadores de condensado", "USD/b", markers.years, condensate_series)),
        ("Gas BN", marker_chart("Marcador de gas BN", "USD/MPC", markers.years, gas_bn_series)),
    ]
    property_charts = [
        distribution_chart(title.split(" (")[0], title.split("(")[-1].rstrip(")"), records)
        for title, records in distributions.items()
    ]
    return Section(
        H3("Visualización de insumos"),
        P(
            "Las gráficas muestran los valores cargados para facilitar la inspección visual; "
            "no aplican umbrales de anomalía no definidos en el procedimiento.",
            cls="muted",
        ),
        Div(*(Div(Span(label, cls="chart-label"), Div(__import__("fasthtml.common", fromlist=["NotStr"]).NotStr(svg)), cls="chart-card") for label, svg in charts), cls="chart-grid"),
        Div(
            *(
                Div(
                    Span("Propiedades · " + title, cls="chart-label"),
                    Div(__import__("fasthtml.common", fromlist=["NotStr"]).NotStr(svg)),
                    cls="chart-card",
                )
                for title, svg in zip(distributions, property_charts)
            ),
            cls="chart-grid",
        ),
        cls="panel",
    )


def _result_table(rows: list[dict[str, object]]) -> object:
    years = range(REPORT_START_YEAR, REPORT_END_YEAR + 1)
    headers = ("Campo", "Hidrocarburo", "Marcador", "API / BTU", "Unidad") + tuple(
        str(year) for year in years
    )
    body = []
    for row in rows:
        values = [
            row["campo"],
            row["tipo_hidrocarburo"],
            row["tipo_gas"],
            f'{float(row["api_btu"]):,.2f}',
            row["unidad"],
            *(
                f'{float(row[f"AÑO {year}"]):,.4f}'
                for year in range(REPORT_START_YEAR, REPORT_END_YEAR + 1)
            ),
        ]
        body.append(Tr(*(Td(value) for value in values)))
    return Div(
        Div(
            Strong(f"{len(rows):,} registros"),
            Span(f"Todos los marcadores · {REPORT_START_YEAR}–{REPORT_END_YEAR}"),
            cls="grid-summary",
        ),
        P(
            "Desplázate verticalmente para recorrer campos y marcadores, y "
            "horizontalmente para consultar todos los años.",
            cls="grid-instructions",
        ),
        Div(
            Table(
                Thead(Tr(*(Th(value, scope="col") for value in headers))),
                Tbody(*body),
                cls="results-table",
            ),
            cls="table-scroll",
        ),
        cls="results-grid",
    )





def _calculation_result(
    *,
    scenario: str,
    markers: object,
    properties: list,
    catalog: set[str],
    rows: list[dict[str, object]],
    run_id: str,
    output_path: Path,
    warnings: list[str],
    property_source: str,
    catalog_source: str,
    calculation_mode: str,
    input_set_version: int,
) -> object:
    distributions = property_distribution(properties, catalog)
    summary = Div(
        Div(Strong(f"{len(rows):,}"), Small(" registros calculados"), cls="metric"),
        Div(Strong(scenario.title()), Small(" escenario"), cls="metric"),
        Div(
            Strong(f"{REPORT_START_YEAR}–{REPORT_END_YEAR}"),
            Small(" horizonte del reporte"),
            cls="metric",
        ),
        cls="metrics",
    )
    metadata = Div(
        Div(Strong("Título"), Span(REPORT_TITLE)),
        Div(Strong("Logo"), Span("GIOEE")),
        Div(Strong("Proceso"), Span("Cartera")),
        Div(
            Strong("Uso del cálculo"),
            Span(CALCULATION_MODES[calculation_mode]),
        ),
        Div(Strong("Versión de insumos"), Span(f"v{input_set_version}")),
        Div(Strong("Fecha de cálculo"), Span(rows[0]["fecha_calculo"])),
        cls="report-meta",
    )
    source_note = P(
        "Propiedades: "
        + property_source
        + " · Catálogo: "
        + catalog_source
        + f" · Set de insumos vigente v{input_set_version}"
        + " · Marcadores: "
        + ", ".join(markers.source_files),
        cls="source-note",
    )
    return Div(
        Section(
            Div(
                H2("Resultado de Cartera"),
                Span("Cálculo generado", cls="status-pill"),
                cls="result-heading",
            ),
            metadata,
            summary,
            Div(
                Strong("Descargar resultados completos"),
                P("Libro Excel con valores anuales y metadatos del reporte."),
                LinkDownload(
                    f"/download/{run_id}",
                    calculation_export_filename(
                        str(rows[0]["fecha_calculo"]),
                        str(rows[0]["escenario"]),
                        input_set_version,
                    ),
                ),
                cls="download-card",
            ),
            source_note,
            _issue_list("Advertencias de calidad", warnings, "warning"),
            cls="panel report-panel",
        ),
        _chart_panel(markers, distributions),
        Section(
            H3("Resultados"),
            _result_table(rows),
            cls="panel",
        ),
        cls="calculation-success",
        data_input_set_version=str(input_set_version),
    )


def LinkDownload(href: str, label: str) -> object:
    from fasthtml.common import A

    return A(label, href=href, download=True, cls="download-link")


def _calculation_error(message: str) -> object:
    return Div(
        H3("No se realizó el cálculo"),
        P(message),
        cls="alert error",
        role="alert",
    )


@rt("/")
def get():
    try:
        current = get_current_input_set()
    except InputStoreError as exc:
        return _page(input_status=str(exc))
    if current is None:
        status = "Aún no hay un set de insumos vigente; se guardará al primer cálculo."
    else:
        status = f"Set de insumos vigente: v{current.version}."
    return _page(input_status=status)


@rt("/healthz")
def get_health():
    return {"status": "ok"}


@rt("/calculate", methods=["POST"])
async def post_calculate(request: Request):
    try:
        form = await request.form()
        scenario = str(form.get("scenario") or "").strip().lower()
        if scenario not in SCENARIO_FOLDERS:
            raise InputError("Selecciona un escenario Alto, Medio o Bajo.")
        calculation_mode = str(
            form.get("calculation_mode") or "reporte_final"
        ).strip().lower()
        if calculation_mode not in CALCULATION_MODES:
            raise InputError("Selecciona Validación o Reporte final de proceso.")

        input_set, warnings = await _load_or_update_input_set(form)
        properties = input_set.properties
        catalog = input_set.catalog
        markers = input_set.markers[scenario]

        run_id = uuid.uuid4().hex
        calculated_at = datetime.now().astimezone().isoformat(timespec="seconds")
        rows = calculate_prices(
            properties,
            catalog,
            markers,
            calculated_at=calculated_at,
            run_id=run_id,
            calculation_mode=calculation_mode,
        )
        if not rows:
            raise InputError("Los insumos no generaron registros de resultados.")
        source_files = [
            input_set.properties_source,
            input_set.catalog_source,
            *input_set.marker_sources[scenario],
        ]
        output_path = write_calculation(
            rows,
            run_id=run_id,
            scenario=scenario.title(),
            calculated_at=calculated_at,
            calculation_mode=calculation_mode,
            input_set_version=input_set.version,
            source_files=source_files,
            warnings=warnings,
        )
        return _calculation_result(
            scenario=scenario,
            markers=markers,
            properties=properties,
            catalog=catalog,
            rows=rows,
            run_id=run_id,
            output_path=output_path,
            warnings=warnings,
            property_source=input_set.properties_source,
            catalog_source=input_set.catalog_source,
            calculation_mode=calculation_mode,
            input_set_version=input_set.version,
        )
    except InputError as exc:
        return _calculation_error(str(exc))
    except InputStoreError as exc:
        return _calculation_error(str(exc))
    except OSError as exc:
        return _calculation_error(f"Error al leer o guardar un archivo: {exc}")


@rt("/download/{run_id}")
def get_download(run_id: str):
    export = calculation_xlsx(run_id)
    if export is None:
        return _calculation_error("No se encontró el reporte solicitado.")
    content, filename = export
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


if __name__ == "__main__":
    from fasthtml.common import serve

    serve(port=4001)
