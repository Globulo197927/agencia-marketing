import re
from pathlib import Path
from PIL import Image
import pytesseract


def extract_text_from_image(image_path: str, lang: str = "spa") -> str:
    img_path = Path(image_path)
    if not img_path.exists():
        raise FileNotFoundError(f"Image not found: {image_path}")

    img = Image.open(img_path)

    if img.width > 3000:
        ratio = 3000 / img.width
        new_size = (3000, int(img.height * ratio))
        img = img.resize(new_size, Image.Resampling.LANCZOS)

    text = pytesseract.image_to_string(img, lang=lang)
    return text


def parse_prices(text: str) -> list[dict]:
    prices = []
    price_pattern = r'(\d+[.,]\d{2})\s*€[^\d]*?(?:/mes|/día|/ud|/línea)?'
    for match in re.finditer(price_pattern, text):
        price_str = match.group(1).replace('.', '').replace(',', '.')
        try:
            price_float = float(price_str)
            prices.append({
                "original": match.group(0),
                "value": price_float,
                "position": match.start()
            })
        except ValueError:
            continue
    return prices


def parse_line_count(text: str) -> dict:
    line_patterns = [
        r'(\d+)\s*líneas?\s*móviles?',
        r'(\d+)\s*lines?\s*mobiles?',
        r'(\d+)\s*líneas?',
        r'(\d+)\s*sedes?',
    ]
    for pattern in line_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return {"lines": int(match.group(1)), "original": match.group(0)}

    return {"lines": 0, "original": ""}


def parse_client_name(text: str) -> dict:
    nif_pattern = r'(?:NIF|CIF|nif|cif)\s*:?\s*([A-Z]\d{8})'
    nif_match = re.search(nif_pattern, text)
    nif = nif_match.group(1) if nif_match else ""

    sa_pattern = r'([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+)*)\s+(?:SA|SL|S\.L\.|S\.A\.)'
    sa_match = re.search(sa_pattern, text)
    company_name = sa_match.group(0) if sa_match else ""

    return {
        "nif": nif,
        "company_name": company_name,
        "raw_text": text[:500]
    }


def parse_product_info(text: str) -> list[dict]:
    products = []

    product_patterns = [
        r'(?:Bundle|Pack|Plan|Tarifa)\s+(?:RED|Infinity|Full Speed|5G|Fibra)\s*(?:PRO|Plus|Pro)?[\s\n]*(\S+(?:\s+\S+)*)',
        r'(?:Fibra|FTTO|Centralita)\s+(?:\d+\s*G?bps|Plus|IP|Cloud)?[\s\n]*(\S+(?:\s+\S+)*)',
    ]

    price_context_pattern = r'(\d+[.,]\d{2})€/mes'
    price_matches = list(re.finditer(price_context_pattern, text))

    for pm in price_matches:
        start = max(0, pm.start() - 100)
        context = text[start:pm.start()]
        product_name = context.strip().split('\n')[-1].strip() if context else ""
        products.append({
            "name": product_name,
            "price": pm.group(0),
            "position": pm.start()
        })

    return products


def extract_crm_data(image_path: str, lang: str = "spa") -> dict:
    text = extract_text_from_image(image_path, lang)

    prices = parse_prices(text)
    line_info = parse_line_count(text)
    client_info = parse_client_name(text)
    products = parse_product_info(text)

    return {
        "raw_text": text,
        "prices": prices,
        "lines": line_info,
        "client": client_info,
        "products": products,
        "confidence": "auto"
    }


def format_crm_data_as_yaml(data: dict) -> str:
    lines = ["# Datos extraídos automáticamente del CRM"]
    lines.append("# Revisa y ajusta los valores antes de generar la presentación")
    lines.append("")

    if data["client"].get("company_name"):
        parts = data["client"]["company_name"].rsplit(" ", 1)
        lines.append(f"client:")
        lines.append(f"  name_line1: \"{parts[0] if len(parts) > 1 else data['client']['company_name']}\"")
        lines.append(f"  name_line2: \"{parts[1] if len(parts) > 1 else ''}\"")
        if data["client"].get("nif"):
            lines.append(f"  nif: \"{data['client']['nif']}\"")

    if data["lines"].get("lines"):
        lines.append(f"  lines: {data['lines']['lines']}")

    if data["prices"]:
        lines.append("")
        lines.append("products:")
        for i, p in enumerate(data["products"][:3]):
            lines.append(f"  - name: \"{p.get('name', f'Producto {i+1}')}\"")
            lines.append(f"    price_monthly: \"{p.get('price', '')}\"")

    lines.append("")
    lines.append("proposal:")
    lines.append(f"  date: \"Mayo 2026\"")

    return "\n".join(lines)