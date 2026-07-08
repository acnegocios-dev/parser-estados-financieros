from __future__ import annotations

import calendar
import re
from dataclasses import asdict, dataclass
from datetime import date
from pathlib import Path


_FILENAME_RE = re.compile(
    r"(?P<rfc>[A-Z&Ñ]{3,4}\d{6}[A-Z0-9]{3})_(?P<year>\d{4})_(?P<month>\d{2})",
    re.IGNORECASE,
)

_MONTHS_ES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}


@dataclass(frozen=True)
class PeriodVariables:
    source_filename: str
    rfc: str
    period_year: int
    period_month: int
    period_ym: str
    period_compact: str
    period_last_day: str
    period_label_bg: str
    period_label_er: str
    period_label_bal: str

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


def extract_period_variables(path: str | Path) -> PeriodVariables:
    source_filename = Path(path).name
    match = _FILENAME_RE.search(source_filename)
    if not match:
        raise ValueError(
            "Could not extract RFC and period from filename. Expected pattern "
            "'<prefix>_<RFC>_<YYYY>_<MM>.<ext>'."
        )

    rfc = match.group("rfc").upper()
    year = int(match.group("year"))
    month = int(match.group("month"))
    if month < 1 or month > 12:
        raise ValueError(f"Invalid period month in filename: {month:02d}")

    last_day = calendar.monthrange(year, month)[1]
    month_name = _MONTHS_ES[month]
    last_date = date(year, month, last_day).isoformat()

    return PeriodVariables(
        source_filename=source_filename,
        rfc=rfc,
        period_year=year,
        period_month=month,
        period_ym=f"{year:04d}-{month:02d}",
        period_compact=f"{year:04d}{month:02d}",
        period_last_day=last_date,
        period_label_bg=f"Al {last_day} de {month_name} de {year:04d}",
        period_label_er=f"Del 1ro de Enero al {last_day} de {month_name} de {year:04d}",
        period_label_bal=f"{month_name.upper()}.{year:04d}",
    )


def period_variables_dict(path: str | Path) -> dict[str, object]:
    return extract_period_variables(path).to_dict()
