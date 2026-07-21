"""Conservative income statement calculation engine.

The engine receives normalized accounting rows and maps accounts to income
statement lines by account code prefixes. It intentionally keeps the mapping
rules editable because the definitive chart of accounts is not available yet.
"""

from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

try:
    from .accounting_profiles import AccountingProfile, MappingRule, load_accounting_profile
except ImportError:  # pragma: no cover - supports PYTHONPATH=src usage.
    from accounting_profiles import AccountingProfile, MappingRule, load_accounting_profile


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


BG_LINE_DEFINITIONS: Tuple[Dict[str, Any], ...] = (
    {"key": "caja_chica", "label": "Caja Chica", "section": "activo", "codes": ("1110",)},
    {"key": "bancos", "label": "Bancos", "section": "activo", "codes": ("1120",)},
    {"key": "inversiones_y_valores", "label": "Inversiones y Valores", "section": "activo", "codes": ("1125",)},
    {"key": "cuentas_por_cobrar", "label": "Cuentas por Cobrar", "section": "activo", "codes": ("1130",)},
    {"key": "contribuciones", "label": "Contribuciones", "section": "activo", "codes": ("1160", "1167", "1190", "1195")},
    {"key": "inventarios", "label": "Inventarios", "section": "activo", "codes": ("1170",)},
    {"key": "deudores_diversos", "label": "Deudores Diversos", "section": "activo", "codes": ("1150",)},
    {"key": "anticipo_a_proveedores", "label": "Anticipo a Proveedores", "section": "activo", "codes": ("1197",)},
    {"key": "pagos_anticipados", "label": "Pagos anticipados", "section": "activo", "codes": ("1196",)},
    {"key": "mobiliario_y_equipo", "label": "Mobiliario y Equipo", "section": "activo", "codes": ("1210",)},
    {"key": "equipo_de_computo", "label": "Equipo de Computo", "section": "activo", "codes": ("1220",)},
    {"key": "equipo_de_transporte", "label": "Equipo de Transporte", "section": "activo", "codes": ("1230",)},
    {"key": "maquinaria_y_equipo", "label": "Maquinaria y Equipo", "section": "activo", "codes": ("1240",)},
    {"key": "depreciacion_acumulada", "label": "Depreciacion acumulada", "section": "activo", "codes": ("1215", "1225", "1235", "1245"), "sign": Decimal("-1")},
    {"key": "proveedores", "label": "Proveedores", "section": "pasivo", "codes": ("2110",)},
    {"key": "acreedores_diversos", "label": "Acreedores Diversos", "section": "pasivo", "codes": ("2120",)},
    {"key": "anticipos_de_clientes", "label": "Anticipos de Clientes", "section": "pasivo", "codes": ("2130",)},
    {"key": "impuestos_por_pagar", "label": "Impuestos por Pagar", "section": "pasivo", "codes": ("2140",)},
    {"key": "otros_pasivos_ptu", "label": "Otros Pasivos PTU", "section": "pasivo", "codes": ("2160",)},
    {"key": "capital_social", "label": "Capital Social", "section": "capital", "codes": ("3100",)},
    {"key": "aportaciones_para_aumentos_de_capital", "label": "Aportaciones para aumentos de capital", "section": "capital", "codes": ("3110",)},
    {"key": "resultados_de_ejercicios_anteriores", "label": "Resultados de ejercicios anteriores", "section": "capital", "codes": ("3160",)},
)


DEFAULT_ACCOUNTING_PROFILE_PATH = (
    Path(__file__).parent / "profiles" / "SME170717GA0-2026-07-v1.json"
)


class ProfileCoverageError(ValueError):
    """Raised when a profile cannot safely produce a financial statement."""

    def __init__(self, blockers: Sequence[Mapping[str, Any]]):
        self.blockers = tuple(dict(item) for item in blockers)
        super().__init__("; ".join(str(item.get("code", "profile_coverage")) for item in blockers))


def load_default_accounting_profile() -> AccountingProfile:
    """Return the sole approved local profile used by this isolated rollout."""

    return load_accounting_profile(DEFAULT_ACCOUNTING_PROFILE_PATH)


def _profile_or_default(profile: AccountingProfile | None) -> AccountingProfile:
    return profile if profile is not None else load_default_accounting_profile()


def _profile_line_definitions(
    profile: AccountingProfile,
    statement: str,
    layout: Sequence[Mapping[str, Any]],
) -> tuple[dict[str, Any], ...]:
    """Project approved profile rules onto generator-owned line geometry.

    Codes live only in the accounting profile.  The static layouts retain
    labels/rows/sections needed to preserve the approved workbook template.
    """

    enabled = set(profile.enabled_lines.get(statement, ()))
    by_line: dict[str, list[MappingRule]] = {}
    for rule in profile.rules:
        if rule.statement == statement and rule.approved:
            by_line.setdefault(rule.line_key, []).append(rule)
    definitions: list[dict[str, Any]] = []
    for item in layout:
        key = str(item["key"])
        if key not in enabled:
            continue
        rules = by_line.get(key, [])
        codes = tuple(rule.account_code for rule in rules if rule.account_code)
        excluded = tuple(
            code for rule in rules for code in rule.exclude_account_codes
        )
        signs = {rule.presentation_sign for rule in rules if rule.account_code}
        if len(signs) > 1:
            raise ProfileCoverageError(({
                "code": "profile_line_sign_ambiguous", "line_key": key,
            },))
        definitions.append({
            **item,
            "codes": codes,
            "exclude_codes": excluded,
            "sign": Decimal(str(next(iter(signs), 1))),
            "rule_ids": tuple(rule.rule_id for rule in rules),
        })
    return tuple(definitions)


def _active_bg_definitions(profile: AccountingProfile | None) -> tuple[dict[str, Any], ...]:
    return _profile_line_definitions(_profile_or_default(profile), "BG", BG_LINE_DEFINITIONS)


def _active_er_definitions(profile: AccountingProfile | None) -> tuple[dict[str, Any], ...]:
    return _profile_line_definitions(_profile_or_default(profile), "ER", ER_MAPPED_LINES)


def build_input_views(rows: Iterable[NormalizedRow]) -> Dict[str, Tuple[Dict[str, Any], ...]]:
    """Preserve BAL's original rows while exposing the leaf-only calculation view."""

    normalized_rows = tuple(_row_as_mapping(row) for row in rows)
    return {
        "all_rows": normalized_rows,
        "calculation_rows": tuple(leaf_account_rows(normalized_rows)),
    }


def build_bal_dataset(rows: Iterable[NormalizedRow]) -> Dict[str, Any]:
    """Build the ordered BAL data without discarding accumulator rows."""

    views = build_input_views(rows)
    all_rows = views["all_rows"]
    codes = [canonical_account_code(row.get("account_code") or row.get("top_account") or "") for row in all_rows]
    bal_rows: List[Dict[str, Any]] = []
    accumulator_rows: List[Dict[str, Any]] = []
    for row, code in zip(all_rows, codes):
        is_accumulator = bool(code and "-" not in code)
        item = {
            "account_raw": row.get("account_raw"),
            "account_code": code,
            "account_name": row.get("account_name"),
            "saldo_inicial": money_to_float(to_decimal(row.get("saldo_inicial"))),
            "debe": money_to_float(to_decimal(row.get("debe"))),
            "haber": money_to_float(to_decimal(row.get("haber"))),
            "saldo_final": money_to_float(to_decimal(row.get("saldo_final"))),
            "parent_code": row.get("parent_code"),
            "nature": row.get("nature"),
            "sat_group_code": row.get("sat_group_code"),
            "catalog_match": row.get("catalog_match"),
            "is_accumulator": is_accumulator,
            "source_row": row.get("source_row"),
        }
        bal_rows.append(item)
        if is_accumulator:
            accumulator_rows.append(item)

    totals = {
        field: money_to_float(sum((to_decimal(row[field]) for row in accumulator_rows), Decimal("0")))
        for field in ("saldo_inicial", "debe", "haber", "saldo_final")
    }
    return {
        "rows": bal_rows,
        "all_rows": bal_rows,
        "calculation_rows": list(views["calculation_rows"]),
        "accumulator_rows": accumulator_rows,
        "sumas_iguales": {
            "accumulator_source_rows": [row["source_row"] for row in accumulator_rows],
            "totals": totals,
        },
    }


def resolve_account_code(
    rows: Iterable[NormalizedRow],
    account_code: str,
    *,
    tolerance: float | Decimal = 1.0,
) -> Dict[str, Any]:
    """Resolve a code once, preferring its accumulator over descendant leaves.

    The returned evidence is intentionally complete so callers can expose an
    ``aggregate_detail_mismatch`` warning without reimplementing this policy.
    """

    views = build_input_views(rows)
    all_rows = views["all_rows"]
    expected = canonical_account_code(account_code)
    exact_rows = tuple(
        row for row in all_rows
        if canonical_account_code(row.get("account_code") or row.get("top_account") or "") == expected
    )
    descendant_leaves, hierarchy_method = _descendant_leaf_rows(
        all_rows, views["calculation_rows"], expected
    )
    accumulator_amount = sum((to_decimal(row.get("saldo_final")) for row in exact_rows), Decimal("0"))
    leaf_amount = sum((to_decimal(row.get("saldo_final")) for row in descendant_leaves), Decimal("0"))
    difference = accumulator_amount - leaf_amount
    tolerance_amount = to_decimal(tolerance)

    if exact_rows:
        amount = accumulator_amount
        policy = "exact_accumulator"
    elif descendant_leaves:
        amount = leaf_amount
        policy = "leaf_fallback"
    else:
        amount = Decimal("0")
        policy = "missing_zero"

    return {
        "account_code": expected,
        "amount": amount,
        "accumulator_amount": accumulator_amount if exact_rows else None,
        "leaf_amount": leaf_amount,
        "difference": difference if exact_rows else None,
        "policy": policy,
        "exact_rows": tuple(exact_rows),
        "leaf_rows": tuple(descendant_leaves),
        "hierarchy_method": hierarchy_method,
        "aggregate_detail_mismatch": bool(exact_rows and descendant_leaves and abs(difference) > tolerance_amount),
    }


def build_bg_dataset(
    rows: Iterable[NormalizedRow],
    *,
    result_ejercicio: Any = None,
    tolerance: float | Decimal = 1.0,
    company: str | None = None,
    period: str | None = None,
    source_path: str | None = None,
    profile: AccountingProfile | None = None,
    enforce_profile_coverage: bool = False,
) -> Dict[str, Any]:
    """Build the BG dataset from account codes, never from manual row numbers."""

    views = build_input_views(rows)
    all_rows = views["all_rows"]
    warnings: List[Dict[str, Any]] = []
    lines: List[Dict[str, Any]] = []
    totals = {"activo": Decimal("0"), "pasivo": Decimal("0"), "capital": Decimal("0")}

    active_profile = _profile_or_default(profile)
    definitions = _active_bg_definitions(active_profile)
    for definition in definitions:
        resolutions = [resolve_account_code(all_rows, code, tolerance=tolerance) for code in definition["codes"]]
        amount = sum((resolution["amount"] for resolution in resolutions), Decimal("0")) * to_decimal(definition.get("sign", 1))
        section = str(definition["section"])
        totals[section] += amount
        for resolution in resolutions:
            warnings.extend(_resolution_warnings(resolution, definition))
        lines.append(
            {
                "key": definition["key"],
                "label": definition["label"],
                "section": section,
                "account_codes": list(definition["codes"]),
                "presentation_sign": int(to_decimal(definition.get("sign", 1))),
                "amount": money_to_float(amount),
                "resolutions": [_resolution_payload(resolution) for resolution in resolutions],
            }
        )

    if result_ejercicio is None:
        result_amount = Decimal("0")
        warnings.append({
            "code": "resultado_ejercicio_no_proporcionado",
            "message": "No se proporciono el resultado generado por ER; se conserva cero.",
        })
    else:
        result_amount = to_decimal(result_ejercicio)
    totals["capital"] += result_amount
    lines.append({
        "key": "resultado_del_ejercicio",
        "label": "Resultado del ejercicio",
        "section": "capital",
        "account_codes": [],
        "amount": money_to_float(result_amount),
        "resolutions": [],
    })

    difference = totals["activo"] - (totals["pasivo"] + totals["capital"])
    tolerance_amount = to_decimal(tolerance)
    balance = {
        "reference": "BG!L47",
        "formula": "F45-L45",
        "total_activo": money_to_float(totals["activo"]),
        "total_pasivo": money_to_float(totals["pasivo"]),
        "capital_contable": money_to_float(totals["capital"]),
        "diferencia_cuadre": float(difference),
        "tolerance": float(tolerance_amount),
        "cuadra": abs(difference) < tolerance_amount,
        "balanza_no_cuadra": abs(difference) >= tolerance_amount,
    }
    coverage = build_profile_coverage(all_rows, active_profile)
    if enforce_profile_coverage:
        require_profile_coverage(coverage)
    return {
        "company": company,
        "period": period,
        "source_path": source_path,
        "lines": lines,
        "warnings": warnings,
        "profile": _profile_dataset_metadata(active_profile),
        "coverage": coverage,
        "effective_rows": _effective_bg_rows(lines),
        "input_views": {"all_rows": len(all_rows), "calculation_rows": len(views["calculation_rows"])},
        "balance": balance,
        **balance,
    }


def build_er_dataset(
    rows: Iterable[NormalizedRow],
    *,
    company: str | None = None,
    period: str | None = None,
    source_path: str | None = None,
    profile: AccountingProfile | None = None,
    enforce_profile_coverage: bool = False,
) -> Dict[str, Any]:
    """Build the deterministic ER dataset from normalized Auditalo rows.

    The accumulated column H is driven by `saldo_final` from leaf accounts.
    Percentages for column J are calculated against H18, matching the manual.
    """

    input_views = build_input_views(rows)
    leaf_rows = input_views["calculation_rows"]
    amounts: Dict[str, Money] = {}
    account_matches: Dict[str, List[Dict[str, Any]]] = {}
    warnings: List[Dict[str, Any]] = []

    active_profile = _profile_or_default(profile)
    definitions = _active_er_definitions(active_profile)
    for definition in definitions:
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
    lines = _er_dataset_lines(amounts, account_matches, definitions=definitions)
    base_amount = amounts.get("ingresos_por_servicios", Decimal("0"))
    unmatched_accounts = _er_unmatched_accounts(leaf_rows, definitions=definitions)
    coverage = build_profile_coverage(input_views["all_rows"], active_profile)
    if enforce_profile_coverage:
        require_profile_coverage(coverage)

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
        "profile": _profile_dataset_metadata(active_profile),
        "coverage": coverage,
        "effective_rows": _effective_er_rows(lines),
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


def _row_as_mapping(row: NormalizedRow | Any) -> Dict[str, Any]:
    if isinstance(row, Mapping):
        return dict(row)
    if hasattr(row, "to_dict"):
        return dict(row.to_dict())
    if hasattr(row, "__dict__"):
        return dict(vars(row))
    raise TypeError(f"Unsupported normalized row: {type(row)!r}")


def leaf_account_rows(rows: Iterable[NormalizedRow]) -> Tuple[NormalizedRow, ...]:
    """Return leaves using catalog parent links before code-prefix fallback."""

    materialized = tuple(_row_as_mapping(row) for row in rows)
    codes = {
        canonical_account_code(row.get("account_code") or row.get("top_account") or "")
        for row in materialized
    }
    children = _catalog_children(materialized)
    leaf_codes: set[str] = set()
    for code in codes:
        if not code:
            continue
        # A catalog parent relation is authoritative for this node.  Prefix
        # matching only fills gaps where the catalog does not provide it.
        if code in children:
            if not children[code]:
                leaf_codes.add(code)
            continue
        if not any(other != code and other.startswith(f"{code}-") for other in codes):
            leaf_codes.add(code)
    return tuple(
        row
        for row in materialized
        if canonical_account_code(row.get("account_code") or row.get("top_account") or "") in leaf_codes
    )


def _catalog_children(rows: Sequence[NormalizedRow]) -> dict[str, set[str]]:
    """Build only valid direct parent evidence; no parent is invented."""

    codes = {
        canonical_account_code(row.get("account_code") or row.get("top_account") or "")
        for row in rows
    }
    children: dict[str, set[str]] = {}
    for row in rows:
        code = canonical_account_code(row.get("account_code") or row.get("top_account") or "")
        parent = canonical_account_code(row.get("parent_code") or "")
        if code and parent and parent in codes and parent != code:
            children.setdefault(parent, set()).add(code)
            children.setdefault(code, set())
    return children


def _descendant_leaf_rows(
    all_rows: Sequence[NormalizedRow],
    calculation_rows: Sequence[NormalizedRow],
    expected: str,
) -> tuple[tuple[NormalizedRow, ...], str]:
    """Resolve descendants via parent_code, then via legacy code prefix."""

    children = _catalog_children(all_rows)
    if expected in children:
        reachable: set[str] = set()
        pending = list(children[expected])
        while pending:
            code = pending.pop()
            if code in reachable:
                continue
            reachable.add(code)
            pending.extend(children.get(code, ()))
        leaves = tuple(
            row for row in calculation_rows
            if canonical_account_code(row.get("account_code") or row.get("top_account") or "") in reachable
        )
        return leaves, "parent_code"
    leaves = tuple(
        row for row in calculation_rows
        if (code := canonical_account_code(row.get("account_code") or row.get("top_account") or ""))
        and code.startswith(f"{expected}-")
    )
    return leaves, "prefix_fallback"


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


def build_profile_coverage(
    rows: Iterable[NormalizedRow], profile: AccountingProfile | None = None
) -> Dict[str, Any]:
    """Account for every material leaf once, with auditable section controls.

    Prefixes are used solely when the catalog does not supply a usable parent
    chain.  The first digit is a control bucket, never a classification rule.
    """

    active_profile = _profile_or_default(profile)
    views = build_input_views(rows)
    leaves = views["calculation_rows"]
    rules = tuple(rule for rule in active_profile.rules if rule.approved and rule.account_code)
    by_code = {
        canonical_account_code(row.get("account_code") or row.get("top_account") or ""): row
        for row in views["all_rows"]
    }
    entries: list[dict[str, Any]] = []
    blockers: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    source_totals = {key: Decimal("0") for key in ("1", "2", "3")}
    assigned_totals = {key: Decimal("0") for key in ("1", "2", "3")}

    for row in leaves:
        code = canonical_account_code(row.get("account_code") or row.get("top_account") or "")
        amount = to_decimal(row.get("saldo_final"))
        section = code[:1] if code[:1] in source_totals else None
        candidates = [rule for rule in rules if _rule_covers_account(code, rule, by_code)]
        line_keys = tuple(sorted({rule.line_key for rule in candidates}))
        status = "assigned" if len(line_keys) == 1 else (
            "unassigned" if not line_keys else "duplicate"
        )
        entry = {
            "account_code": code,
            "source_row": row.get("source_row"),
            "amount": money_to_float(amount),
            "section": section,
            "status": status,
            "line_keys": list(line_keys),
            "rule_ids": [rule.rule_id for rule in candidates],
        }
        entries.append(entry)
        # Section controls compare presentation amounts.  A contra-asset's
        # source balance and its displayed sign are therefore reconciled on
        # the same basis; an unmapped account remains raw and blocks anyway.
        presentation_amount = amount * to_decimal(candidates[0].presentation_sign) if status == "assigned" else amount
        if section:
            source_totals[section] += presentation_amount
        if status == "assigned":
            sign = candidates[0].presentation_sign
            if section:
                assigned_totals[section] += amount * to_decimal(sign)
        elif amount == 0:
            warnings.append({
                "code": "profile_mapping_missing_zero", "severity": "warning", **entry,
            })
        else:
            blockers.append({
                "code": "profile_mapping_unassigned_material" if status == "unassigned" else "profile_mapping_duplicate",
                "severity": "blocking", **entry,
            })

    section_controls: dict[str, dict[str, Any]] = {}
    for section in ("1", "2", "3"):
        difference = source_totals[section] - assigned_totals[section]
        control = {
            "source_amount": money_to_float(source_totals[section]),
            "assigned_amount": money_to_float(assigned_totals[section]),
            "difference": money_to_float(difference),
            "status": "blocking" if abs(difference) >= Decimal("1") else "ok",
        }
        section_controls[section] = control
        if control["status"] == "blocking":
            blockers.append({
                "code": "profile_section_coverage_difference",
                "severity": "blocking",
                "section": section,
                **control,
            })
    return {
        "profile_id": active_profile.profile_id,
        "profile_version": active_profile.profile_version,
        "entries": entries,
        "assigned": sum(entry["status"] == "assigned" for entry in entries),
        "unassigned": sum(entry["status"] == "unassigned" for entry in entries),
        "ambiguous": sum(entry["status"] == "duplicate" for entry in entries),
        "duplicates": sum(entry["status"] == "duplicate" for entry in entries),
        "section_controls": section_controls,
        "blockers": blockers,
        "warnings": warnings,
    }


def require_profile_coverage(coverage: Mapping[str, Any]) -> None:
    """Block generation on material gaps, duplicate assignment, or >= $1 drift."""

    blockers = coverage.get("blockers", ())
    if blockers:
        raise ProfileCoverageError(blockers)


def _rule_covers_account(
    account_code: str,
    rule: MappingRule,
    by_code: Mapping[str, NormalizedRow],
) -> bool:
    expected = canonical_account_code(rule.account_code or "")
    if not expected or _code_matches(account_code, tuple(canonical_account_code(code) for code in rule.exclude_account_codes)):
        return False
    if account_code == expected:
        return True
    # Catalog hierarchy is primary.  A prefix may only stand in for missing
    # parent evidence, preserving legacy input behavior without inventing it.
    current = by_code.get(account_code)
    seen: set[str] = set()
    while current is not None:
        parent = canonical_account_code(current.get("parent_code") or "")
        if not parent or parent in seen:
            break
        if parent == expected:
            return True
        seen.add(parent)
        current = by_code.get(parent)
    return account_code.startswith(f"{expected}-")


def _profile_dataset_metadata(profile: AccountingProfile) -> Dict[str, Any]:
    return {
        "accounting_profile_id": profile.profile_id,
        "accounting_profile_version": profile.profile_version,
        "accounting_profile_status": profile.status,
        "accounting_profile_company": profile.company_name,
        "accounting_profile_rfc": profile.rfc,
        "accounting_profile_valid_from": profile.valid_from.isoformat(),
        "accounting_profile_valid_to": None if profile.valid_to is None else profile.valid_to.isoformat(),
        "rfc": profile.rfc,
        "base_taxonomy_version": profile.taxonomy_version,
        "catalog_source_sha256": profile.catalog_identity.source_sha256,
        "catalog_semantic_sha256": profile.catalog_identity.semantic_sha256,
        "generator_profile_id": None if profile.generator_profile is None else profile.generator_profile.profile_id,
        "generator_profile_version": None if profile.generator_profile is None else profile.generator_profile.profile_version,
    }


def _effective_bg_rows(lines: Sequence[Mapping[str, Any]]) -> Dict[str, list[str]]:
    return {
        "mapped_line_keys": [str(line["key"]) for line in lines],
        "subtotal_line_keys": ["total_circulante", "total_no_circulante", "total_otros_activos", "total_activo", "total_pasivo", "total_capital", "cuadre"],
    }


def _effective_er_rows(lines: Sequence[Mapping[str, Any]]) -> Dict[str, list[Any]]:
    return {
        "line_keys": [str(line["key"]) for line in lines],
        "rows": [int(line["excel_row"]) for line in lines],
        "subtotal_line_keys": [
            str(line["key"]) for line in lines if line.get("line_type") == "calculated"
        ],
    }


def _resolution_payload(resolution: Mapping[str, Any]) -> Dict[str, Any]:
    return {
        "account_code": resolution["account_code"],
        "amount": money_to_float(to_decimal(resolution["amount"])),
        "accumulator_amount": (
            None if resolution["accumulator_amount"] is None
            else money_to_float(to_decimal(resolution["accumulator_amount"]))
        ),
        "leaf_amount": money_to_float(to_decimal(resolution["leaf_amount"])),
        "difference": (
            None if resolution["difference"] is None
            else float(to_decimal(resolution["difference"]))
        ),
        "policy": resolution["policy"],
        "hierarchy_method": resolution.get("hierarchy_method"),
        "exact_source_rows": [row.get("source_row") for row in resolution.get("exact_rows", ())],
        "leaf_source_rows": [row.get("source_row") for row in resolution.get("leaf_rows", ())],
    }


def _resolution_warnings(
    resolution: Mapping[str, Any],
    definition: Mapping[str, Any],
) -> List[Dict[str, Any]]:
    if resolution["policy"] == "missing_zero":
        return [{
            "code": "cuenta_no_encontrada",
            "line_key": definition["key"],
            "label": definition["label"],
            "account_code": resolution["account_code"],
            "message": f"No se encontro la cuenta {resolution['account_code']}; se devuelve cero.",
        }]
    if resolution["aggregate_detail_mismatch"]:
        return [{
            "code": "aggregate_detail_mismatch",
            "line_key": definition["key"],
            "account_code": resolution["account_code"],
            "accumulator": money_to_float(to_decimal(resolution["accumulator_amount"])),
            "leaf_sum": money_to_float(to_decimal(resolution["leaf_amount"])),
            "difference": float(to_decimal(resolution["difference"])),
            "policy": resolution["policy"],
            "message": (
                f"El acumulador {resolution['account_code']} difiere de la suma de hojas; "
                "se uso el acumulador exacto."
            ),
        }]
    return []


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
    *,
    definitions: Sequence[Mapping[str, Any]] | None = None,
) -> List[Dict[str, Any]]:
    line_definitions: List[Mapping[str, Any]] = []
    line_definitions.extend(definitions if definitions is not None else _active_er_definitions(None))
    line_definitions.extend(ER_ZERO_LINES)
    line_definitions.extend(ER_CALCULATED_LINES)
    line_definitions.sort(key=lambda item: int(item["excel_row"]))
    mapped_signs = {
        str(item["key"]): int(to_decimal(item.get("sign", 1)))
        for item in (definitions if definitions is not None else _active_er_definitions(None))
    }

    base_amount = amounts.get("ingresos_por_servicios", Decimal("0"))
    calculated_keys = {str(line["key"]) for line in ER_CALCULATED_LINES}
    lines: List[Dict[str, Any]] = []
    for definition in line_definitions:
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
                "presentation_sign": mapped_signs.get(key, 1),
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
        "parent_code": row.get("parent_code"),
        "nature": row.get("nature"),
        "sat_group_code": row.get("sat_group_code"),
        "catalog_match": row.get("catalog_match"),
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


def _er_unmatched_accounts(
    rows: Sequence[NormalizedRow], *, definitions: Sequence[Mapping[str, Any]] | None = None
) -> List[Dict[str, Any]]:
    definitions = [
        definition for definition in (definitions if definitions is not None else _active_er_definitions(None))
        if definition.get("codes")
    ]
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
