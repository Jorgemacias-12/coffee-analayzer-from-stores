from src.config import CONFIG
from src.load import load_excel_files
from src.filter import filter_products
from src.date import enrich_dates
from src.report import export_reports


def main():
    data = (
        load_excel_files()
    )

    filtered = (
        filter_products(
            data
        )
    )

    if filtered.empty:
        print(
            "Sin cafés."
        )

        return

    enriched = (
        enrich_dates(
            filtered
        )
    )

    export_reports(
        enriched
    )

    print(
        "Proceso completado."
    )


if __name__ == "__main__":
    main()
