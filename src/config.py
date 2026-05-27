from __future__ import annotations
from dataclasses import dataclass


@dataclass
class ReportConfig:

    input_folder = "ventas"

    output_folder = "reportes"

    date_column = "Fecha"

    product_column = "Descripción"

    store_column = "Establecimiento"

    product_filter = (
        "FRATELLI"
    )

    excluded_keywords = (
        "TAPA",
        "VASO",
        "RELLENO"
    )

    hour_interval = 2


CONFIG = ReportConfig()
