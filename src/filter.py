from src.config import CONFIG
from src.date import normalize_dates


def filter_products(
    df,
):

    result = (
        df.copy()
    )

    if (
        CONFIG.date_column
        not in result.columns
    ):

        raise ValueError(
            f"No existe {CONFIG.date_column}"
        )

    result[
        CONFIG.date_column
    ] = (
        normalize_dates(
            result[
                CONFIG.date_column
            ]
        )
    )

    result = (
        result.loc[
            result[
                CONFIG.date_column
            ]
            .notna()
        ]
        .copy()
    )

    if result.empty:

        raise ValueError(
            "Sin fechas válidas."
        )

    description = (
        result[
            CONFIG.product_column
        ]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(
            r'\s+',
            ' ',
            regex=True
        )
    )

    include = (
        description
        .str.startswith(
            CONFIG.product_filter,
            na=False,
        )
    )

    exclude = (
        description
        .str.contains(
            "|".join(
                CONFIG.excluded_keywords
            ),
            case=False,
            na=False,
        )
    )

    result = (
        result[
            include
            &
            ~exclude
        ]
        .copy()
    )

    result[
        CONFIG.product_column
    ] = (
        result[
            CONFIG.product_column
        ]
        .astype(str)
        .str.strip()
        .str.replace(
            r'\s+',
            ' ',
            regex=True
        )
    )

    result = (
        result.drop_duplicates()
        .reset_index(drop=True)
    )

    return result
