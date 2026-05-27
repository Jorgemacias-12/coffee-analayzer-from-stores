from pathlib import Path
from src.config import CONFIG

import pandas as pd


def load_excel_files():
    folder = Path(
        CONFIG.input_folder
    )

    if not folder.exists():
        raise FileNotFoundError(
            "Carpeta no encontrada."
        )

    files = [
        *folder.glob("*.xls"),
        *folder.glob("*.xlsx")
    ]

    if not files:
        raise ValueError(
            "No hay excels."
        )

    frames = []

    for file in files:

        print(f"Leyendo {file.name}")

        df = pd.read_excel(
            file,
            dtype=object
        )

        frames.append(
            df
        )

    return pd.concat(
        frames,
        ignore_index=True
    )
