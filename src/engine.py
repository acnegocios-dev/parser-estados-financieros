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
