from pathlib import Path
import pandas as pd
from src.config import CONFIG

from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment
)
from openpyxl.chart import BarChart, Reference


def build_store_report(
    df,
):

    grouped = (

        df.groupby(
            [
                "Fecha",
                "Hora",
                CONFIG.product_column,
            ],
            as_index=False,
        )

        .size()

        .rename(
            columns={
                CONFIG.product_column:
                    "Producto",

                "size":
                    "Cant Venta",
            }
        )
    )

    rows = []

    dates = sorted(
        grouped[
            "Fecha"
        ]
        .unique()
    )

    for date in dates:

        day = (

            grouped[
                grouped[
                    "Fecha"
                ]
                ==
                date
            ]

            .copy()

            .sort_values(
                [
                    "Hora",
                    "Cant Venta",
                    "Producto",
                ],
                ascending=[
                    True,
                    False,
                    True,
                ],
            )
        )

        rows.append(
            day
        )

        interval_resume = (

            day.groupby(
                "Hora"
            )

            [
                "Cant Venta"
            ]

            .sum()

            .sort_values(
                ascending=False
            )
        )

        if (
            interval_resume.empty
        ):

            continue

        best_hour = (
            interval_resume
            .index[0]
        )

        rows.append(

            pd.DataFrame(
                [
                    {
                        "Producto":
                            "INTERVALO MÁS VENDIDO",

                        "Fecha":
                            "",

                        "Hora":
                            best_hour,
                    }
                ]
            )
        )

    return pd.concat(
        rows,
        ignore_index=True,
    )


def sanitize_filename(
    value,
):

    invalid = (
        '<>:"/\\|?*'
    )

    for char in invalid:

        value = (
            value.replace(
                char,
                "_",
            )
        )

    return value


def export_reports(
    df,
):

    folder = Path(
        CONFIG.output_folder
    )

    folder.mkdir(
        exist_ok=True
    )

    stores = (
        df[
            CONFIG.store_column
        ]
        .dropna()
        .unique()
    )

    for store in stores:

        print(
            f"Generando {store}"
        )

        report = (
            build_store_report(
                df[
                    df[
                        CONFIG.store_column
                    ]
                    ==
                    store
                ]
            )
        )

        report = report[
            [
                "Producto",
                "Fecha",
                "Hora",
                "Cant Venta",
            ]
        ]

        path = (
            folder
            /
            (
                sanitize_filename(
                    str(
                        store
                    )
                )
                +
                ".xlsx"
            )
        )

        with pd.ExcelWriter(
            path,
            engine="openpyxl",
        ) as writer:

            report.to_excel(
                writer,
                index=False,
            )

            ws = (
                writer.sheets[
                    "Sheet1"
                ]
            )

            fill = (
                PatternFill(
                    start_color="FFF2CC",
                    end_color="FFF2CC",
                    fill_type="solid",
                )
            )

            font = (
                Font(
                    bold=True,
                )
            )

            align = (
                Alignment(
                    horizontal="center",
                    vertical="center",
                )
            )

            # --- NUEVA LÓGICA PARA GRÁFICAS Y DISEÑO ---

            # Guardaremos los límites de inicio y fin de cada día para graficar individualmente
            day_blocks = []
            start_row = 2
            current_date_str = ""

            for row in range(
                2,
                ws.max_row + 1,
            ):
                # Detectar la fecha actual del bloque
                val_fecha = ws.cell(row, 2).value
                if val_fecha and val_fecha != "":
                    current_date_str = str(val_fecha)

                if (
                    ws.cell(
                        row,
                        1,
                    ).value
                    !=
                    "INTERVALO MÁS VENDIDO"
                ):
                    continue

                # Al llegar a "INTERVALO MÁS VENDIDO", sabemos que el bloque del día termina en la fila anterior
                end_row = row - 1
                if end_row >= start_row:
                    day_blocks.append((start_row, end_row, current_date_str))

                interval = (
                    ws.cell(
                        row,
                        3,
                    ).value
                )

                # limpiar fila de datos crudos
                for col in range(
                    1,
                    5,
                ):
                    ws.cell(
                        row,
                        col,
                    ).value = None

                # combinar
                ws.merge_cells(
                    f"A{row}:B{row}"
                )

                ws.merge_cells(
                    f"C{row}:D{row}"
                )

                ws[
                    f"A{row}"
                ] = (
                    "INTERVALO MÁS VENDIDO"
                )

                ws[
                    f"C{row}"
                ] = (
                    interval
                )

                for col in (
                    "A",
                    "C",
                ):

                    cell = (
                        ws[
                            f"{col}{row}"
                        ]
                    )

                    cell.fill = (
                        fill
                    )

                    cell.font = (
                        font
                    )

                    cell.alignment = (
                        align
                    )

                # El siguiente bloque de día empezará dos filas abajo (saltando este resumen)
                start_row = row + 1

            # --- GENERACIÓN DE GRÁFICAS POR BLOQUE DE DÍA ---
            # Colocaremos las gráficas a la derecha de la tabla (Columna F)
            chart_placement_row = 2

            for b_start, b_end, b_date in day_blocks:
                chart = BarChart()
                chart.type = "col"
                chart.style = 10
                chart.title = f"Ventas por Producto - {b_date}"
                chart.y_axis.title = "Cantidad"
                chart.x_axis.title = "Productos"
                chart.height = 12  # Tamaño adaptado para que se vea bien
                chart.width = 18
                chart.legend = None  # Quitamos la leyenda porque solo hay 1 serie

                # Datos numéricos: Columna D ("Cant Venta") desde b_start hasta b_end
                data_ref = Reference(
                    ws, min_col=4, min_row=b_start - 1, max_row=b_end)
                # Categorías (Eje X): Columna A ("Producto")
                cats_ref = Reference(
                    ws, min_col=1, min_row=b_start, max_row=b_end)

                chart.add_data(data_ref, titles_from_data=True)
                chart.set_categories(cats_ref)

                # Añadir la gráfica en la columna F, alineada con su bloque de datos
                ws.add_chart(chart, f"F{chart_placement_row}")

                # Desplazar la posición de la siguiente gráfica hacia abajo para que no se encimen
                chart_placement_row += 25

            # Ajuste de ancho de columnas automático
            for column in ws.columns:

                try:
                    # Evitar que tome en cuenta celdas combinadas largas para calcular el ancho
                    width = max(
                        len(str(cell.value or ""))
                        for cell in column
                        if cell.coordinate not in ws.merged_cells
                    )

                    ws.column_dimensions[
                        column[
                            0
                        ]
                        .column_letter
                    ].width = (
                        max(width, 10)
                        +
                        4
                    )

                except:
                    pass
