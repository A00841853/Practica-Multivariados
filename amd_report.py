"""Exportacion portable de los mismos resultados a CSV y Excel.

El XLSX contiene resultados ya simulados, no vuelve a simular al editar celdas.
Para cambiar supuestos, ejecutar nuevamente amd_monte_carlo.py.
"""
import csv
from pathlib import Path
from tempfile import TemporaryDirectory


def export_reports(rows, csv_path, xlsx_path):
    """Prepara ambos archivos antes de reemplazar las versiones anteriores."""
    try:
        import xlsxwriter
    except ImportError as exc:
        raise RuntimeError(
            "Falta XlsxWriter. Ejecuta: python -m pip install -r requirements.txt"
        ) from exc
    csv_path, xlsx_path = Path(csv_path), Path(xlsx_path)
    if csv_path.resolve() == xlsx_path.resolve():
        raise ValueError("Las rutas de salida deben ser distintas.")
    # Temporales en cada destino permiten reemplazos atomicos por archivo.
    # No es una transaccion entre archivos: si Excel bloquea el segundo destino,
    # el CSV puede haberse actualizado. El mensaje de error pide repetir.
    with TemporaryDirectory(dir=csv_path.parent) as csv_temp, \
            TemporaryDirectory(dir=xlsx_path.parent) as excel_temp:
        staged_csv = Path(csv_temp) / "resumen.csv"
        staged_xlsx = Path(excel_temp) / "analisis.xlsx"
        with staged_csv.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
            writer.writeheader()
            writer.writerows(rows)
        write_excel(rows, staged_xlsx, xlsxwriter)
        try:
            staged_csv.replace(csv_path)
            staged_xlsx.replace(xlsx_path)
        except PermissionError as exc:
            raise RuntimeError(
                "Cierra los archivos en Excel y vuelve a ejecutar. "
                "El CSV puede haberse actualizado antes del bloqueo del XLSX."
            ) from exc


def write_excel(rows, path, xlsxwriter):
    """Dos hojas y cuatro graficas editables, vinculadas a celdas numericas."""
    colors = ["#2563EB", "#B45309", "#15803D", "#7C3AED", "#BE185D", "#475569"]
    money_code = '"$"#,##0.00;("$"#,##0.00)'
    keys = list(rows[0])
    index = {key: i for i, key in enumerate(keys)}
    ordered = sorted(enumerate(rows), key=lambda pair: (
        pair[1]["annual_return"], pair[1]["months"]
    ))
    horizons = sorted({row["months"] for row in rows})
    returns = sorted({row["annual_return"] for row in rows})
    longest = max(horizons)

    with xlsxwriter.Workbook(path) as book:
        book.set_properties({"title": "AMD: escenarios Monte Carlo",
                             "comments": "Ejercicio educativo. No es recomendacion financiera."})
        summary = book.add_worksheet("Resumen")
        data = book.add_worksheet("Datos completos")
        base = {"font_name": "Arial", "font_size": 10, "valign": "vcenter"}
        normal = book.add_format(base)
        title = book.add_format({**base, "font_size": 16, "bold": True,
                                 "font_color": "#17365D"})
        note = book.add_format({**base, "font_color": "#475569"})
        header = book.add_format({**base, "bold": True, "bg_color": "#17365D",
                                  "font_color": "white", "text_wrap": True,
                                  "align": "center"})
        money = book.add_format({**base, "num_format": money_code})
        percent = book.add_format({**base, "num_format": "0.00%"})
        integer = book.add_format({**base, "num_format": "#,##0"})
        for sheet in (summary, data):
            sheet.hide_gridlines(2)
            sheet.set_default_row(22)
            sheet.set_zoom(85)
            sheet.set_landscape()
            sheet.fit_to_pages(1, 0)
        summary.set_tab_color("#17365D")
        summary.set_column("A:A", 18, percent)
        summary.set_column("B:B", 9, integer)
        summary.set_column("C:F", 15, money)
        summary.set_column("G:J", 17, percent)
        summary.write("A2", "AMD: escenarios Monte Carlo", title)
        summary.write("A3", f"Precio inicial: USD {rows[0]['s0']:,.2f}. "
                      f"Volatilidad anual: {rows[0]['volatility']:.2%}. "
                      f"Simulaciones por fila: {rows[0]['simulations']:,}. "
                      f"Semilla: {rows[0]['seed']}.", note)
        summary.write("A4", "Cada fila es un supuesto de rendimiento anual, "
                      "no una probabilidad de que ese escenario ocurra.", note)
        summary.write("A5", "Precios en USD. Probabilidades al final del plazo. "
                      "P05-P95 contiene el 90% central simulado.", note)
        summary.write("A6", "Resultados guardados: cambia supuestos ejecutando Python de nuevo. "
                      "Ejercicio educativo; no es recomendacion financiera.", note)

        # La hoja completa conserva exactamente las columnas y precision del CSV.
        data.write("A2", "Datos completos de la simulacion", title)
        data.write("A3", "Fuente: parametros proporcionados al script y simulacion GBM. "
                   "No se consultaron cotizaciones ni volatilidad de mercado.", note)
        data.write("A4", "mean: promedio simulado. theoretical_mean: promedio exacto del modelo. "
                   "median = p50. pXX: XX% de resultados queda debajo de ese precio.", note)
        data.write("A5", "prob_loss_gt_20pct: precio final < 0.8 x S0. "
                   "No mide la caida maxima durante el periodo.", note)
        data_columns = []
        for col, key in enumerate(keys):
            fmt = (percent if key in ("annual_return", "volatility")
                   or key.startswith("prob_") else
                   integer if key in ("months", "simulations", "seed") else money)
            data.set_column(col, col, 24 if key.startswith("prob_") else 19, fmt)
            data_columns.append({"header": key, "format": fmt, "header_format": header})
        data.add_table(6, 0, 6 + len(rows), len(keys) - 1, {
            "name": "DatosSimulados", "style": "Table Style Medium 2",
            "columns": data_columns, "data": [[row[k] for k in keys] for row in rows],
        })
        data.set_row(6, 40)
        data.freeze_panes(7, 2)

        fields = [
            ("annual_return", "Rendimiento anual", percent),
            ("months", "Meses", integer),
            ("mean", "Media (USD)", money), ("median", "Mediana (USD)", money),
            ("p05", "P05 (USD)", money), ("p95", "P95 (USD)", money),
            ("prob_st_gt_s0", "P(ganar)", percent),
            ("prob_st_gt_700", "P(precio > 700)", percent),
            ("prob_st_lt_500", "P(precio < 500)", percent),
            ("prob_loss_gt_20pct", "P(perder > 20%)", percent),
        ]
        summary.add_table(7, 0, 7 + len(rows), 9, {
            "name": "ResumenEscenarios", "style": "Table Style Medium 2",
            "columns": [{"header": label, "format": fmt, "header_format": header}
                        for _, label, fmt in fields],
        })
        summary.set_row(7, 40)
        summary.freeze_panes(8, 2)
        for offset, (original, row) in enumerate(ordered, 8):
            for col, (key, _, fmt) in enumerate(fields):
                ref = xlsxwriter.utility.xl_rowcol_to_cell(original + 7, index[key])
                summary.write_formula(offset, col, f"='Datos completos'!{ref}", fmt, row[key])
        summary.conditional_format(8, 9, 7 + len(rows), 9, {
            "type": "2_color_scale", "min_type": "num", "min_value": 0,
            "max_type": "num", "max_value": 1, "min_color": "#FFFFFF",
            "max_color": "#FCA5A5",
        })

        # Bloques de apoyo visibles debajo de los datos, para auditar las graficas.
        support_start = len(rows) + 10
        data.write(support_start, 0, f"Comparacion a {longest} meses", header)
        compare = [row for _, row in ordered if row["months"] == longest]
        compare_fields = ["annual_return", "mean", "median", "p05", "p95"]
        data.write_row(support_start + 1, 0, ["Escenario anual", "Media", "Mediana", "P05", "P95"], header)
        for offset, row in enumerate(compare, support_start + 2):
            source = next(i for i, r in enumerate(rows) if r is row) + 7
            data.write(offset, 0, f"{row['annual_return']:+.2%}", normal)
            for col, key in enumerate(compare_fields[1:], 1):
                ref = xlsxwriter.utility.xl_rowcol_to_cell(source, index[key])
                data.write_formula(offset, col, f"={ref}", money, row[key])

        def style_chart(chart, name, probability=False):
            chart.set_title({"name": name, "name_font": {"name": "Arial", "size": 12}})
            chart.set_legend({"position": "bottom", "font": {"name": "Arial", "size": 10}})
            chart.set_y_axis({"name": "Probabilidad" if probability else "Precio (USD)",
                              "num_format": "0%" if probability else '"$"#,##0',
                              "min": 0, **({"max": 1} if probability else {})})
            chart.set_size({"width": 525, "height": 330})
            chart.set_chartarea({"border": {"none": True}})

        start, end = support_start + 2, support_start + 1 + len(compare)
        chart = book.add_chart({"type": "column"})
        for col, name, color in [(1, "Media", colors[0]), (2, "Mediana", colors[2])]:
            chart.add_series({"name": name, "categories": [data.name, start, 0, end, 0],
                              "values": [data.name, start, col, end, col],
                              "fill": {"color": color}, "border": {"none": True}})
        style_chart(chart, f"Media y mediana a {longest} meses")
        chart_row = len(rows) + 11
        summary.insert_chart(chart_row, 0, chart)

        interval = book.add_chart({"type": "column"})
        for col, name, color in [(3, "P05", "#94A3B8"), (2, "Mediana", colors[0]),
                                 (4, "P95", "#60A5FA")]:
            interval.add_series({"name": name, "categories": [data.name, start, 0, end, 0],
                                 "values": [data.name, start, col, end, col],
                                 "fill": {"color": color}, "border": {"none": True}})
        style_chart(interval, f"Percentiles de precio a {longest} meses")
        summary.insert_chart(chart_row, 5, interval)

        next_block = end + 4
        for metric, chart_name, anchor_col in [
            ("prob_st_gt_700", "Probabilidad de terminar arriba de USD 700", 0),
            ("prob_loss_gt_20pct", "Probabilidad de perder mas de 20%", 5),
        ]:
            data.write(next_block, 0, chart_name, normal)
            data.write(next_block + 1, 0, "Meses", header)
            chart = book.add_chart({"type": "scatter", "subtype": "straight_with_markers"})
            for col, annual_return in enumerate(returns, 1):
                data.write(next_block + 1, col, f"{annual_return:+.2%} anual", header)
                for offset, horizon in enumerate(horizons, next_block + 2):
                    data.write(offset, 0, horizon, integer)
                    original, row = next((i, r) for i, r in enumerate(rows)
                                         if r["annual_return"] == annual_return and r["months"] == horizon)
                    ref = xlsxwriter.utility.xl_rowcol_to_cell(original + 7, index[metric])
                    data.write_formula(offset, col, f"={ref}", percent, row[metric])
                chart.add_series({
                    "name": [data.name, next_block + 1, col],
                    "categories": [data.name, next_block + 2, 0, next_block + 1 + len(horizons), 0],
                    "values": [data.name, next_block + 2, col, next_block + 1 + len(horizons), col],
                    "line": {"color": colors[(col - 1) % len(colors)], "width": 2},
                    "marker": {"type": "circle", "size": 5},
                })
            style_chart(chart, chart_name, probability=True)
            chart.set_x_axis({"name": "Horizonte (meses)", "min": 0,
                              "max": longest + 1, "num_format": "0"})
            summary.insert_chart(chart_row + 17, anchor_col, chart)
            next_block += len(horizons) + 5
        summary.print_area(0, 0, chart_row + 34, 9)
        data.print_area(0, 0, next_block, len(keys) - 1)
