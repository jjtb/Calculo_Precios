from __future__ import annotations

import html
from statistics import median
from typing import Mapping, Sequence

from pricing import MarkerSet, OIL_REFERENCE_FIELDS, oil_price


PALETTE = (
    "#52784f",
    "#b58b37",
    "#668f83",
    "#986f47",
    "#7d8f56",
    "#4275a1",
    "#965d7c",
)


def _scale(value: float, low: float, high: float, start: float, end: float) -> float:
    if high == low:
        return (start + end) / 2
    return start + (value - low) * (end - start) / (high - low)


def oil_marker_series(markers: MarkerSet) -> dict[str, list[float]]:
    series = {
        name: [values[year] for year in markers.years]
        for name, values in markers.oil.items()
    }
    for field, api in OIL_REFERENCE_FIELDS.items():
        series[f"{field} ({api:g}° API)"] = [
            oil_price(api, markers, year) for year in markers.years
        ]
    return series


def marker_chart(
    title: str,
    unit: str,
    years: Sequence[int],
    series: Mapping[str, Sequence[float]],
) -> str:
    width = 720
    left, right, top = 54, 18, 30
    plot_height = 150
    legend_columns = 3
    legend_rows = (len(series) + legend_columns - 1) // legend_columns
    year_label_y = top + plot_height + 22
    legend_start_y = year_label_y + 22
    legend_step = 15
    height = legend_start_y + max(legend_rows, 1) * legend_step + 8
    values = [value for entries in series.values() for value in entries]
    if not values or not years:
        return '<p class="muted">Sin datos suficientes para graficar.</p>'
    y_min, y_max = min(values), max(values)
    if y_min == y_max:
        padding = abs(y_min) * 0.05 or 1
        y_min -= padding
        y_max += padding
    plot_width = width - left - right
    svg = [
        f'<svg class="chart" viewBox="0 0 {width} {height}" role="img" aria-label="{html.escape(title)}">',
        f'<text class="chart-title" x="{left}" y="18">{html.escape(title)} ({html.escape(unit)})</text>',
    ]
    for tick in range(5):
        value = y_min + (y_max - y_min) * tick / 4
        y = top + plot_height * (4 - tick) / 4
        svg.append(f'<line class="grid-line" x1="{left}" y1="{y:.1f}" x2="{width-right}" y2="{y:.1f}"/>')
        svg.append(f'<text class="axis-label" x="{left-8}" y="{y+4:.1f}" text-anchor="end">{value:.2f}</text>')
    for year in (years[0], years[len(years) // 2], years[-1]):
        x = _scale(year, years[0], years[-1] or years[0] + 1, left, width - right)
        svg.append(f'<text class="axis-label" x="{x:.1f}" y="{year_label_y}" text-anchor="middle">{year}</text>')
    for index, (label, prices) in enumerate(series.items()):
        points = []
        for year, price in zip(years, prices):
            x = _scale(year, years[0], years[-1] or years[0] + 1, left, width - right)
            y = _scale(price, y_min, y_max, top + plot_height, top)
            points.append(f"{x:.1f},{y:.1f}")
        color = PALETTE[index % len(PALETTE)]
        svg.append(f'<polyline class="series-line" stroke="{color}" points="{" ".join(points)}"/>')
        legend_x = left + (index % legend_columns) * 210
        legend_y = legend_start_y + (index // legend_columns) * legend_step
        svg.append(f'<line x1="{legend_x}" y1="{legend_y-4}" x2="{legend_x+18}" y2="{legend_y-4}" stroke="{color}" stroke-width="3"/>')
        svg.append(f'<text class="legend-label" x="{legend_x+24}" y="{legend_y}">{html.escape(label)}</text>')
    svg.append("</svg>")
    return "".join(svg)


def distribution_chart(
    title: str, unit: str, records: Sequence[tuple[str, float]]
) -> str:
    width, height = 720, 175
    left, right, top, bottom = 52, 18, 28, 36
    finite = sorted(value for _, value in records if value is not None)
    if not finite:
        return '<p class="muted">Sin datos suficientes para graficar.</p>'
    low, high = finite[0], finite[-1]
    if low == high:
        delta = abs(low) * 0.05 or 1
        low -= delta
        high += delta
    plot_width = width - left - right
    plot_height = height - top - bottom
    points = []
    for index, value in enumerate(finite):
        x = left + (plot_width * index / max(len(finite) - 1, 1))
        y = _scale(value, low, high, top + plot_height, top)
        points.append(f"{x:.1f},{y:.1f}")
    middle = median(finite)
    svg = (
        f'<svg class="chart" viewBox="0 0 {width} {height}" role="img" aria-label="{html.escape(title)}">'
        f'<text class="chart-title" x="{left}" y="18">{html.escape(title)} ({html.escape(unit)})</text>'
        f'<line class="grid-line" x1="{left}" y1="{top+plot_height}" x2="{width-right}" y2="{top+plot_height}"/>'
        f'<polyline class="distribution-line" points="{" ".join(points)}"/>'
        f'<text class="axis-label" x="{left}" y="{height-8}">mín. {finite[0]:.2f}</text>'
        f'<text class="axis-label" x="{width/2}" y="{height-8}" text-anchor="middle">mediana {middle:.2f}</text>'
        f'<text class="axis-label" x="{width-right}" y="{height-8}" text-anchor="end">máx. {finite[-1]:.2f}</text>'
        "</svg>"
    )
    return svg
