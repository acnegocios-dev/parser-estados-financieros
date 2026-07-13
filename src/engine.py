"""Conservative income statement calculation engine.

The engine receives normalized accounting rows and maps accounts to income
statement lines by account code prefixes. It intentionally keeps the mapping
rules editable because the definitive chart of accounts is not available yet.
"""

from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple


Money = Decimal
NormalizedRow = Mapping[str, Any]


# Editable mapping rules. Rules are evaluated from top to bottom, so more
# specific prefixes must appear before broader prefixes.
MAP_RULES: Tuple[Dict[str, Any], ...] = (
    {
        "prefixes": ("4",),
        "line_key": "ventas_ingresos_netos",
        "label": "Ventas / ingresos netos",
        "normal_balance": "credit",
        "category": "income",
        "description": "Ingresos ordinarios. Se presentan positivos cuando el haber supera al debe.",
    },
    {
        "prefixes": ("5",),
        "line_key": "costo_ventas",
        "label": "Costo de ventas",
        "normal_balance": "debit",
        "category": "cost",
        "description": "Costos directos. Se presentan positivos cuando el debe supera al haber.",
    },
    {
        "prefixes": ("6",),
        "line_key": "gastos_operacion",
        "label": "Gastos de operacion",
        "normal_balance": "debit",
        "category": "expense",
        "description": "Gastos de administracion, venta y operacion.",
    },
    {
        "prefixes": ("700", "701", "702", "71"),
        "line_key": "otros_ingresos",
        "label": "Otros ingresos",
        "normal_balance": "credit",
        "category": "other_income",
        "description": "Otros productos o ingresos no ordinarios, si el catalogo usa clase 7.",
    },
    {
        "prefixes": ("703", "704", "705", "72"),
        "line_key": "otros_gastos",
        "label": "Otros gastos",
        "normal_balance": "debit",
        "category": "other_expense",
        "description": "Otros gastos no ordinarios, si el catalogo usa clase 7.",
    },
    {
        "prefixes": ("730", "731", "732", "74"),
        "line_key": "productos_financieros",
        "label": "Productos financieros",
        "normal_balance": "credit",
        "category": "financial_income",
        "description": "Productos financieros. Regla provisional por prefijo.",
    },
    {
        "prefixes": ("733", "734", "735", "75", "76"),
        "line_key": "gastos_financieros",
        "label": "Gastos financieros",
        "normal_balance": "debit",
        "category": "financial_expense",
        "description": "Gastos financieros. Regla provisional por prefijo.",
    },
    {
        "prefixes": ("77", "78", "79"),
        "line_key": "impuestos",
        "label": "Impuestos a la utilidad",
        "normal_balance": "debit",
        "category": "tax",
        "description": "Impuestos del periodo, cuando el catalogo los ubica en clase 7.",
    },
    {
        "prefixes": ("7",),
        "line_key": "otros_gastos",
        "label": "Otros gastos",
        "normal_balance": "debit",
        "category": "other_expense",
        "description": "Regla conservadora de cierre para cuentas clase 7 no mas especificas.",
    },
)


STATEMENT_TEMPLATE: Tuple[Dict[str, Any], ...] = (
    {"key": "ventas_ingresos_netos", "label": "Ventas / ingresos netos", "type": "mapped", "excel_row": 18},
    {"key": "costo_ventas", "label": "Costo de ventas", "type": "mapped"},
    {"key": "utilidad_bruta", "label": "Utilidad bruta", "type": "calculated"},
    {"key": "gastos_operacion", "label": "Gastos de operacion", "type": "mapped"},
    {"key": "resultado_operacion", "label": "Resultado de operacion", "type": "calculated"},
    {"key": "otros_ingresos", "label": "Otros ingresos", "type": "mapped"},
    {"key": "otros_gastos", "label": "Otros gastos", "type": "mapped"},
    {"key": "productos_financieros", "label": "Productos financieros", "type": "mapped"},
    {"key": "gastos_financieros", "label": "Gastos financieros", "type": "mapped"},
    {"key": "resultado_antes_impuestos", "label": "Resultado antes de impuestos", "type": "calculated"},
    {"key": "impuestos", "label": "Impuestos a la utilidad", "type": "mapped"},
    {"key": "resultado_ejercicio", "label": "Resultado del ejercicio", "type": "calculated"},
)


FORMULAS: Dict[str, str] = {
    "utilidad_bruta": "ventas_ingresos_netos - costo_ventas",
    "resultado_operacion": "utilidad_bruta - gastos_operacion",
    "resultado_antes_impuestos": (
        "resultado_operacion + otros_ingresos - otros_gastos "
        "+ productos_financieros - gastos_financieros"
    ),
    "resultado_ejercicio": "resultado_antes_impuestos - impuestos",
}


# ER layout and account-code mapping from the Portal CI Obsidian notes. These
# rules intentionally use account codes only, never Excel source row numbers.
ER_MAPPED_LINES: Tuple[Dict[str, Any], ...] = (
    {
        "key": "ingresos_por_servicios",
        "label": "Ingresos por servicios",
        "excel_row": 18,
        "codes": ("4110",),
        "exclude_codes": ("4110-9999",),
    },
    {
        "key": "costo_de_ventas",
        "label": "Costo de ventas",
        "excel_row": 23,
        "codes": ("5110", "5115"),
    },
    {"key": "sueldos_y_salarios", "label": "Sueldos y salarios", "excel_row": 28, "codes": ("6110",)},
    {"key": "impuestos_y_derechos", "label": "Impuestos y derechos", "excel_row": 29, "codes": ("6120",)},
    {"key": "honorarios", "label": "Honorarios", "excel_row": 30, "codes": ("6125",)},
    {"key": "arrendamiento", "label": "Arrendamiento", "excel_row": 31, "codes": ("6130",)},
    {"key": "seguros_y_fianzas", "label": "Seguros y fianzas", "excel_row": 32, "codes": ("6135",)},
    {"key": "servicios", "label": "Servicios", "excel_row": 33, "codes": ("6140",)},
    {"key": "capacitacion_al_personal", "label": "Capacitacion al personal", "excel_row": 34, "codes": ("6144",)},
    {"key": "fletes_y_o_mensajeria", "label": "Fletes y/o mensajeria", "excel_row": 35, "codes": ("6145", "6146")},
    {"key": "seguridad_e_higiene", "label": "Seguridad e higiene", "excel_row": 36, "codes": ("6147",)},
    {"key": "mantenimiento", "label": "Mantenimiento", "excel_row": 37, "codes": ("6150",)},
    {"key": "combustibles", "label": "Combustibles", "excel_row": 38, "codes": ("6160",)},
    {"key": "propaganda_y_publicidad", "label": "Propaganda y publicidad", "excel_row": 39, "codes": ("6155",)},
    {"key": "cuotas_y_suscripciones", "label": "Cuotas y suscripciones", "excel_row": 40, "codes": ("6170",)},
    {"key": "gastos_de_viaje", "label": "Gastos de viaje", "excel_row": 41, "codes": ("6175",)},
    {"key": "herrajes_y_herramientas", "label": "Herrajes y herramientas", "excel_row": 42, "codes": ("6177",)},
    {"key": "papeleria_y_art_de_oficina", "label": "Papeleria y art. de oficina", "excel_row": 43, "codes": ("6165",)},
    {"key": "depreciaciones", "label": "Depreciaciones", "excel_row": 44, "codes": ("6185",)},
    {"key": "recargos", "label": "Recargos", "excel_row": 45, "codes": ("6190",)},
    {
        "key": "varios",
        "label": "Varios",
        "excel_row": 46,
        "codes": ("6148", "6176", "6195"),
    },
    {"key": "uniformes", "label": "Uniformes", "excel_row": 47, "codes": ("6196",)},
    {"key": "no_deducibles", "label": "No deducibles", "excel_row": 50, "codes": ("6290",)},
    {"key": "otros_productos", "label": "Otros productos", "excel_row": 56, "codes": ("4110-9999",)},
    {"key": "otros_gastos", "label": "Otros gastos", "excel_row": 57, "codes": ()},
    {"key": "productos_financieros", "label": "Productos financieros", "excel_row": 61, "codes": ("7510",)},
    {
        "key": "gastos_financieros",
        "label": "Gastos financieros",
        "excel_row": 62,
        "codes": ("6410",),
        "sign": Decimal("-1"),
    },
    {"key": "isr_del_ejercicio", "label": "ISR del ejercicio", "excel_row": 67, "codes": ("6510-0001",)},
    {"key": "ptu_del_ejercicio", "label": "PTU del ejercicio", "excel_row": 68, "codes": ("6510-0002",)},
)


ER_CALCULATED_LINES: Tuple[Dict[str, Any], ...] = (
    {
        "key": "ingresos_netos",
        "label": "Ingresos netos",
        "excel_row": 20,
        "formula": "ingresos_por_servicios + descuentos_o_bonificaciones",
    },
    {
        "key": "utilidad_bruta",
        "label": "Utilidad bruta",
        "excel_row": 25,
        "formula": "ingresos_netos - costo_de_ventas",
    },
    {
        "key": "gastos_de_operacion",
        "label": "Gastos de operacion",
        "excel_row": 51,
        "formula": "sum(gastos_operacion_detalle)",
    },
    {
        "key": "utilidad_perdida_operacion",
        "label": "Utilidad o perdida de operacion",
        "excel_row": 53,
        "formula": "utilidad_bruta - gastos_de_operacion",
    },
    {
        "key": "total_otros_ingresos",
        "label": "Total otros ingresos",
        "excel_row": 58,
        "formula": "otros_productos + otros_gastos",
    },
    {
        "key": "resultado_integral_financiamiento",
        "label": "Resultado integral de financiamiento",
        "excel_row": 63,
        "formula": "productos_financieros + gastos_financieros",
    },
    {
        "key": "resultado_antes_impuestos",
        "label": "Resultado antes de impuestos",
        "excel_row": 65,
        "formula": "utilidad_perdida_operacion + total_otros_ingresos + resultado_integral_financiamiento",
    },
    {
        "key": "resultado_ejercicio",
        "label": "Resultado del ejercicio",
        "excel_row": 70,
        "formula": "resultado_antes_impuestos - isr_del_ejercicio - ptu_del_ejercicio",
    },
)


ER_ZERO_LINES: Tuple[Dict[str, Any], ...] = (
    {
        "key": "descuentos_o_bonificaciones",
        "label": "Descuentos o bonificaciones",
        "excel_row": 19,
        "formula": "Sin cuenta definida en mapeo inicial; se conserva cero.",
    },
)


ER_EXPENSE_DETAIL_KEYS = (
    "sueldos_y_salarios",
    "impuestos_y_derechos",
    "honorarios",
    "arrendamiento",
    "seguros_y_fianzas",
    "servicios",
    "capacitacion_al_personal",
    "fletes_y_o_mensajeria",
    "seguridad_e_higiene",
    "mantenimiento",
    "combustibles",
    "propaganda_y_publicidad",
    "cuotas_y_suscripciones",
    "gastos_de_viaje",
    "herrajes_y_herramientas",
    "papeleria_y_art_de_oficina",
    "depreciaciones",
    "recargos",
    "varios",
    "uniformes",
    "no_deducibles",
)


def build_er_dataset(
    rows: Iterable[NormalizedRow],
    *,
    company: str | None = None,
    period: str | None = None,
    source_path: str | None = None,
) -> Dict[str, Any]:
    """Build the deterministic ER dataset from normalized Auditalo rows.

    The accumulated column H is driven by `saldo_final` from leaf accounts.
    Percentages for column J are calculated against H18, matching the manual.
    """

    leaf_rows = leaf_account_rows(rows)
    amounts: Dict[str, Money] = {}
    account_matches: Dict[str, List[Dict[str, Any]]] = {}
    warnings: List[Dict[str, Any]] = []

    for definition in ER_MAPPED_LINES:
        key = str(definition["key"])
        matches = _matching_rows(leaf_rows, definition)
        sign = to_decimal(definition.get("sign", Decimal("1")))
        amount = sum((to_decimal(row.get("saldo_final")) for row in matches), Decimal("0")) * sign
        amounts[key] = amount
        account_matches[key] = [_matched_account_row(row) for row in matches]

        if definition.get("codes") and not matches:
            warnings.append(_missing_account_warning(definition))

    for definition in ER_ZERO_LINES:
        amounts[str(definition["key"])] = Decimal("0")
        account_matches[str(definition["key"])] = []

    _calculate_er_totals(amounts)
    lines = _er_dataset_lines(amounts, account_matches)
    base_amount = amounts.get("ingresos_por_servicios", Decimal("0"))
    unmatched_accounts = _er_unmatched_accounts(leaf_rows)

    return {
        "company": company,
        "period": period,
        "source_path": source_path,
        "lines": lines,
        "statement_lines": lines,
        "accounting_dataset": _er_accounting_dataset(account_matches),
        "unmatched_accounts": unmatched_accounts,
        "missing_accounts": [
            warning["account_code"]
            for warning in warnings
            if warning.get("code") == "cuenta_no_encontrada"
        ],
        "warnings": warnings,
        "formulas": {str(line["key"]): str(line["formula"]) for line in ER_CALCULATED_LINES},
        "raw_amounts": {key: str(value) for key, value in amounts.items()},
        "base_line": {
            "line_key": "ingresos_por_servicios",
            "excel_cell": "H18",
            "amount": money_to_float(base_amount),
            "percentage_column": "J",
        },
        "sign_policy": {
            "source_amount": "saldo_final acumulado de cuentas hoja",
            "operating_expenses": "positivo como rubro",
            "financial_expenses": "saldo_final de 6410 multiplicado por -1",
            "results": "pueden ser negativos",
        },
        "source_rows_used": len(leaf_rows),
    }


def build_income_statement(
    rows: Iterable[NormalizedRow],
    rules: Sequence[Mapping[str, Any]] = MAP_RULES,
) -> Dict[str, Any]:
    """Build the income statement dataset from normalized accounting rows.

    Args:
        rows: Iterable of normalized rows with source_row, account_raw,
            account_code, account_name, top_account, saldo_inicial, debe,
            haber and saldo_final.
        rules: Optional editable mapping rules. Defaults to MAP_RULES.

    Returns:
        A dictionary containing statement lines, accounting dataset,
        formulas/equivalent calculations and unmatched accounts.
    """

    totals = _zero_totals()
    accounting_dataset: List[Dict[str, Any]] = []
    unmatched_accounts: List[Dict[str, Any]] = []

    for row in rows:
        account_code = normalize_account_code(row.get("account_code") or row.get("top_account") or "")
        rule = classify_account(account_code, rules)

        if rule is None:
            unmatched_accounts.append(_unmatched_account(row, account_code))
            continue

        amount = calculate_account_amount(row, rule["normal_balance"])
        line_key = str(rule["line_key"])
        totals[line_key] = totals.get(line_key, Decimal("0")) + amount
        accounting_dataset.append(_dataset_row(row, account_code, rule, amount))

    calculated = calculate_statement_totals(totals)
    totals.update(calculated)

    statement_lines = _statement_lines(totals)
    base_amount = totals["ventas_ingresos_netos"]
    _attach_percentages(statement_lines, base_amount)

    return {
        "statement_lines": statement_lines,
        "accounting_dataset": accounting_dataset,
        "unmatched_accounts": unmatched_accounts,
        "formulas": dict(FORMULAS),
        "base_line": {
            "line_key": "ventas_ingresos_netos",
            "excel_cell": "H18",
            "amount": money_to_float(base_amount),
            "percentage_column": "J",
        },
        "sign_policy": {
            "credit": "haber - debe; fallback saldo_final as reported if only balance is present",
            "debit": "debe - haber; fallback saldo_final as reported if only balance is present",
        },
    }


def classify_account(
    account_code: Any,
    rules: Sequence[Mapping[str, Any]] = MAP_RULES,
) -> Optional[Mapping[str, Any]]:
    """Return the first mapping rule that matches an account code prefix."""

    normalized = normalize_account_code(account_code)
    if not normalized:
        return None

    for rule in rules:
        prefixes = tuple(str(prefix) for prefix in rule.get("prefixes", ()))
        if normalized.startswith(prefixes):
            return rule
    return None


def calculate_account_amount(row: NormalizedRow, normal_balance: str) -> Money:
    """Calculate a signed presentation amount for an account row.

    For result accounts, movements are preferred over ending balance because the
    income statement is period based. If both debe and haber are zero/missing,
    saldo_final is used as the presentation value observed in Auditalo.
    """

    debe = to_decimal(row.get("debe"))
    haber = to_decimal(row.get("haber"))
    saldo_final = to_decimal(row.get("saldo_final"))

    has_movement = debe != 0 or haber != 0
    balance = str(normal_balance).strip().lower()

    if balance == "credit":
        return haber - debe if has_movement else saldo_final
    if balance == "debit":
        return debe - haber if has_movement else saldo_final

    raise ValueError(f"Unsupported normal_balance: {normal_balance!r}")


def calculate_statement_totals(totals: Mapping[str, Money]) -> Dict[str, Money]:
    """Calculate derived income statement lines."""

    ventas = totals.get("ventas_ingresos_netos", Decimal("0"))
    costo = totals.get("costo_ventas", Decimal("0"))
    gastos_operacion = totals.get("gastos_operacion", Decimal("0"))
    otros_ingresos = totals.get("otros_ingresos", Decimal("0"))
    otros_gastos = totals.get("otros_gastos", Decimal("0"))
    productos_financieros = totals.get("productos_financieros", Decimal("0"))
    gastos_financieros = totals.get("gastos_financieros", Decimal("0"))
    impuestos = totals.get("impuestos", Decimal("0"))

    utilidad_bruta = ventas - costo
    resultado_operacion = utilidad_bruta - gastos_operacion
    resultado_antes_impuestos = (
        resultado_operacion
        + otros_ingresos
        - otros_gastos
        + productos_financieros
        - gastos_financieros
    )
    resultado_ejercicio = resultado_antes_impuestos - impuestos

    return {
        "utilidad_bruta": utilidad_bruta,
        "resultado_operacion": resultado_operacion,
        "resultado_antes_impuestos": resultado_antes_impuestos,
        "resultado_ejercicio": resultado_ejercicio,
    }


def normalize_account_code(value: Any) -> str:
    """Normalize account codes to digits for prefix matching."""

    text = "" if value is None else str(value)
    return "".join(character for character in text if character.isdigit())


def to_decimal(value: Any) -> Money:
    """Convert accounting values to Decimal without adding dependencies."""

    if value is None or value == "":
        return Decimal("0")
    if isinstance(value, Decimal):
        return value
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        return Decimal(str(value))

    text = str(value).strip()
    if not text:
        return Decimal("0")

    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]

    cleaned = (
        text.replace("$", "")
        .replace(",", "")
        .replace(" ", "")
        .replace("\u00a0", "")
    )
    amount = Decimal(cleaned)
    return -amount if negative else amount


def money_to_float(value: Money) -> float:
    """Round money values for JSON-friendly output."""

    return float(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def percent_to_float(value: Money) -> float:
    """Round percentage values for JSON-friendly output."""

    return float(value.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP))


def _zero_totals() -> Dict[str, Money]:
    return {
        line["key"]: Decimal("0")
        for line in STATEMENT_TEMPLATE
        if line["type"] == "mapped"
    }


def _dataset_row(
    row: NormalizedRow,
    account_code: str,
    rule: Mapping[str, Any],
    amount: Money,
) -> Dict[str, Any]:
    return {
        "source_row": row.get("source_row"),
        "account_raw": row.get("account_raw"),
        "account_code": account_code,
        "account_name": row.get("account_name"),
        "top_account": row.get("top_account"),
        "line_key": rule["line_key"],
        "line_label": rule["label"],
        "normal_balance": rule["normal_balance"],
        "amount": money_to_float(amount),
        "raw": {
            "saldo_inicial": money_to_float(to_decimal(row.get("saldo_inicial"))),
            "debe": money_to_float(to_decimal(row.get("debe"))),
            "haber": money_to_float(to_decimal(row.get("haber"))),
            "saldo_final": money_to_float(to_decimal(row.get("saldo_final"))),
        },
    }


def _unmatched_account(row: NormalizedRow, account_code: str) -> Dict[str, Any]:
    return {
        "source_row": row.get("source_row"),
        "account_raw": row.get("account_raw"),
        "account_code": account_code,
        "account_name": row.get("account_name"),
        "top_account": row.get("top_account"),
        "saldo_final": money_to_float(to_decimal(row.get("saldo_final"))),
        "reason": "No matching account-code prefix in MAP_RULES.",
    }


def _statement_lines(totals: Mapping[str, Money]) -> List[Dict[str, Any]]:
    lines: List[Dict[str, Any]] = []
    for template in STATEMENT_TEMPLATE:
        key = str(template["key"])
        excel_row = template.get("excel_row")
        lines.append(
            {
                "line_key": key,
                "label": template["label"],
                "line_type": template["type"],
                "amount_column": "H",
                "amount_cell": f"H{excel_row}" if excel_row else None,
                "amount": money_to_float(totals.get(key, Decimal("0"))),
                "formula": FORMULAS.get(key),
            }
        )
    return lines


def _attach_percentages(lines: List[Dict[str, Any]], base_amount: Money) -> None:
    for line in lines:
        amount = to_decimal(line["amount"])
        percentage = Decimal("0") if base_amount == 0 else amount / base_amount
        line["percentage_column"] = "J"
        line["percentage_of"] = "H18"
        line["percentage"] = percent_to_float(percentage)


def canonical_account_code(value: Any) -> str:
    """Normalize account codes while preserving hyphenated subaccounts."""

    text = "" if value is None else str(value).strip().upper()
    clean = []
    previous_was_separator = False
    for character in text:
        if character.isalnum():
            clean.append(character)
            previous_was_separator = False
        elif character in ("-", " ", "."):
            if clean and not previous_was_separator:
                clean.append("-")
                previous_was_separator = True
    return "".join(clean).strip("-")


def leaf_account_rows(rows: Iterable[NormalizedRow]) -> Tuple[NormalizedRow, ...]:
    """Return leaf/detail accounts so accumulator rows are not double counted."""

    materialized = tuple(rows)
    codes = {
        canonical_account_code(row.get("account_code") or row.get("top_account") or "")
        for row in materialized
    }
    leaf_codes = {
        code
        for code in codes
        if code and not any(other != code and other.startswith(f"{code}-") for other in codes)
    }
    return tuple(
        row
        for row in materialized
        if canonical_account_code(row.get("account_code") or row.get("top_account") or "") in leaf_codes
    )


def _matching_rows(
    rows: Sequence[NormalizedRow],
    definition: Mapping[str, Any],
) -> Tuple[NormalizedRow, ...]:
    codes = tuple(canonical_account_code(code) for code in definition.get("codes", ()))
    excluded = tuple(canonical_account_code(code) for code in definition.get("exclude_codes", ()))
    if not codes:
        return ()
    return tuple(
        row
        for row in rows
        if _code_matches(canonical_account_code(row.get("account_code") or row.get("top_account") or ""), codes)
        and not _code_matches(canonical_account_code(row.get("account_code") or row.get("top_account") or ""), excluded)
    )


def _code_matches(account_code: str, expected_codes: Sequence[str]) -> bool:
    return any(
        account_code == expected or account_code.startswith(f"{expected}-")
        for expected in expected_codes
        if expected
    )


def _calculate_er_totals(amounts: Dict[str, Money]) -> None:
    amounts["ingresos_netos"] = (
        amounts.get("ingresos_por_servicios", Decimal("0"))
        + amounts.get("descuentos_o_bonificaciones", Decimal("0"))
    )
    amounts["utilidad_bruta"] = amounts["ingresos_netos"] - amounts.get("costo_de_ventas", Decimal("0"))
    amounts["gastos_de_operacion"] = sum(
        (amounts.get(key, Decimal("0")) for key in ER_EXPENSE_DETAIL_KEYS),
        Decimal("0"),
    )
    amounts["utilidad_perdida_operacion"] = (
        amounts["utilidad_bruta"] - amounts["gastos_de_operacion"]
    )
    amounts["total_otros_ingresos"] = (
        amounts.get("otros_productos", Decimal("0")) + amounts.get("otros_gastos", Decimal("0"))
    )
    amounts["resultado_integral_financiamiento"] = (
        amounts.get("productos_financieros", Decimal("0"))
        + amounts.get("gastos_financieros", Decimal("0"))
    )
    amounts["resultado_antes_impuestos"] = (
        amounts["utilidad_perdida_operacion"]
        + amounts["total_otros_ingresos"]
        + amounts["resultado_integral_financiamiento"]
    )
    amounts["resultado_ejercicio"] = (
        amounts["resultado_antes_impuestos"]
        - amounts.get("isr_del_ejercicio", Decimal("0"))
        - amounts.get("ptu_del_ejercicio", Decimal("0"))
    )


def _er_dataset_lines(
    amounts: Mapping[str, Money],
    account_matches: Mapping[str, Sequence[Mapping[str, Any]]],
) -> List[Dict[str, Any]]:
    definitions: List[Mapping[str, Any]] = []
    definitions.extend(ER_MAPPED_LINES)
    definitions.extend(ER_ZERO_LINES)
    definitions.extend(ER_CALCULATED_LINES)
    definitions.sort(key=lambda item: int(item["excel_row"]))

    base_amount = amounts.get("ingresos_por_servicios", Decimal("0"))
    calculated_keys = {str(line["key"]) for line in ER_CALCULATED_LINES}
    lines: List[Dict[str, Any]] = []
    for definition in definitions:
        key = str(definition["key"])
        amount = amounts.get(key, Decimal("0"))
        percentage = Decimal("0") if base_amount == 0 else amount / base_amount
        matches = list(account_matches.get(key, ()))
        lines.append(
            {
                "key": key,
                "line_key": key,
                "label": definition["label"],
                "excel_row": int(definition["excel_row"]),
                "amount_column": "H",
                "amount_cell": f"H{int(definition['excel_row'])}",
                "period_amount": money_to_float(amount),
                "accumulated_amount": money_to_float(amount),
                "amount": money_to_float(amount),
                "percentage_column": "J",
                "percentage_of": "H18",
                "percentage": percent_to_float(percentage),
                "accounts": matches,
                "line_type": "calculated" if key in calculated_keys else "mapped",
            }
        )
    return lines


def _matched_account_row(row: NormalizedRow) -> Dict[str, Any]:
    return {
        "source_row": row.get("source_row"),
        "account_raw": row.get("account_raw"),
        "account_code": canonical_account_code(row.get("account_code") or row.get("top_account") or ""),
        "account_name": row.get("account_name"),
        "top_account": row.get("top_account"),
        "saldo_final": money_to_float(to_decimal(row.get("saldo_final"))),
    }


def _missing_account_warning(definition: Mapping[str, Any]) -> Dict[str, Any]:
    return {
        "code": "cuenta_no_encontrada",
        "line_key": definition["key"],
        "label": definition["label"],
        "account_code": ", ".join(str(code) for code in definition.get("codes", ())),
        "message": f"No se encontro cuenta esperada para {definition['label']}; se devuelve cero.",
    }


def _er_accounting_dataset(
    account_matches: Mapping[str, Sequence[Mapping[str, Any]]],
) -> List[Dict[str, Any]]:
    dataset: List[Dict[str, Any]] = []
    for line_key, rows in account_matches.items():
        for row in rows:
            item = dict(row)
            item["line_key"] = line_key
            dataset.append(item)
    return dataset


def _er_unmatched_accounts(rows: Sequence[NormalizedRow]) -> List[Dict[str, Any]]:
    definitions = [definition for definition in ER_MAPPED_LINES if definition.get("codes")]
    unmatched: List[Dict[str, Any]] = []
    for row in rows:
        account_code = canonical_account_code(row.get("account_code") or row.get("top_account") or "")
        if any(
            _code_matches(account_code, tuple(canonical_account_code(code) for code in definition.get("codes", ())))
            and not _code_matches(
                account_code,
                tuple(canonical_account_code(code) for code in definition.get("exclude_codes", ())),
            )
            for definition in definitions
        ):
            continue
        unmatched.append(_unmatched_account(row, account_code))
    return unmatched
