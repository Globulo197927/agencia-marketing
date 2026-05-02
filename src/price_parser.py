import re
from typing import Optional


def parse_price_string(price_str: str) -> Optional[float]:
    if not price_str:
        return None
    clean = price_str.replace("€", "").replace("/mes", "").replace("/día", "").replace("/ud", "").replace(" ", "")
    clean = clean.replace(".", "").replace(",", ".")
    try:
        return float(clean)
    except ValueError:
        return None


def format_price(value: float) -> str:
    if value == int(value):
        return f"{int(value):,}".replace(",", ".")
    formatted = f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return formatted


def calculate_vodafone_pricing(data: dict) -> dict:
    client = data.get("client", {})
    products = data.get("products", [])
    economic = data.get("economic", {})
    lines = client.get("lines", 1)

    bundle_price = parse_price_string(products[0].get("price_monthly", "0")) if len(products) > 0 else 0
    fibra_price = parse_price_string(products[1].get("price_monthly", "0")) if len(products) > 1 else 0
    centralita_price = parse_price_string(products[2].get("price_monthly", "0")) if len(products) > 2 else 0

    total_per_line = bundle_price or 0
    total_monthly = total_per_line * lines + fibra_price * (1 if fibra_price > 0 else 0) + centralita_price * (1 if centralita_price > 0 else 0)

    return {
        "total_monthly": f"{format_price(total_monthly)} €/mes",
        "total_per_line": f"{format_price(total_per_line)} €/línea/mes",
        "lines": lines,
    }


def calculate_lynks_pricing(data: dict) -> dict:
    client = data.get("client", {})
    proposal = data.get("proposal", {})
    options = data.get("options", [])

    sedes = client.get("sedes", 0)
    extensiones = proposal.get("extensiones", 0)
    canales = proposal.get("canales", 0)
    ddi_unit_price = parse_price_string(proposal.get("ddi_unit_price", "1,50€/mes"))
    terminal_unit_price = parse_price_string(proposal.get("terminal_unit_price", "8,50€/ud/mes"))

    ddi_total = sedes * ddi_unit_price if ddi_unit_price else 0
    terminal_total = extensiones * terminal_unit_price if terminal_unit_price else 0

    calculated_options = []
    for opt in options:
        ftto_unit_price = parse_price_string(opt.get("ftto_unit_price", "42,00€/mes"))
        ftto_total = sedes * ftto_unit_price if ftto_unit_price else 0
        total = ftto_total + ddi_total + terminal_total + 0

        calculated_options.append({
            **opt,
            "ftto_total": f"{format_price(ftto_total)}€",
            "ddi_total": f"{format_price(ddi_total)}€",
            "terminal_total": f"{format_price(terminal_total)}€",
            "total_price": f"{format_price(total)} €/mes",
        })

    return {
        "options": calculated_options,
        "ddi_total": f"{format_price(ddi_total)}€/mes",
        "terminal_total": f"{format_price(terminal_total)}€/mes",
    }


def merge_ocr_with_manual(ocr_data: dict, manual_data: dict) -> dict:
    merged = {}
    for key in set(list(ocr_data.keys()) + list(manual_data.keys())):
        if key in manual_data:
            merged[key] = manual_data[key]
        elif key in ocr_data:
            merged[key] = ocr_data[key]
    return merged