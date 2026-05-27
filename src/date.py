import pandas as pd
from src.config import CONFIG


def normalize_dates(values: pd.Series):
    values = (
        values.replace(
            {
                "": None,
                "nan": None,
                "NaT": None
            }
        )
    )

    normal = pd.to_datetime(
        values,
        errors="coerce",
        dayfirst=True,
    )

    missing = (
        normal.isna()
    )

    numeric = pd.to_numeric(
        values.where(
            missing
        ),
        errors="coerce",
    )

    numeric = numeric.where(
        numeric.between(
            1,
            70000,
        )
    )

    excel = pd.Series(
        pd.NaT,
        index=values.index,
        dtype="datetime64[ns]",
    )

    valid = (
        numeric.notna()
    )

    if valid.any():

        excel.loc[
            valid
        ] = pd.to_datetime(
            numeric.loc[
                valid
            ],
            unit="D",
            origin="1899-12-30",
            errors="coerce",
        )

    result = (
        normal.fillna(
            excel
        )
    )

    return pd.to_datetime(
        result,
        errors="coerce",
    )


def format_hour(
    hour,
):

    suffix = (
        "am"
        if hour < 12
        else "pm"
    )

    hour = (
        hour % 12
    )

    if hour == 0:

        hour = 12

    return (
        f"{hour}{suffix}"
    )


def build_time_range(
    hour,
):

    start = (
        hour
        //
        CONFIG.hour_interval
    ) * (
        CONFIG.hour_interval
    )

    end = min(
        start
        +
        CONFIG.hour_interval,
        24,
    )

    start_formatted = format_hour(start)
    end_formatted = (
        format_hour(end if end < 24 else 0)
    )

    return (
        f"{start_formatted}"
        f" - "
        f"{end_formatted}"
    )


def enrich_dates(
    df,
):

    result = (
        df.copy()
    )

    dates = pd.to_datetime(
        result[
            CONFIG.date_column
        ],
        errors="coerce",
    )

    valid = (
        dates.notna()
    )

    result = (
        result[
            valid
        ]
        .copy()
    )

    dates = (
        dates[
            valid
        ]
    )

    if result.empty:

        raise ValueError(
            "No quedaron fechas válidas."
        )

    result[
        "Fecha"
    ] = (
        dates.dt.date
    )

    result[
        "Hora"
    ] = (
        dates.dt.hour
        .astype(int)
        .apply(
            build_time_range
        )
    )

    return result
