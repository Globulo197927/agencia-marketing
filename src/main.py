#!/usr/bin/env python3
import argparse
import sys
import json
import yaml
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.config_loader import (
    load_client_data,
    merge_overrides,
    dict_to_vodafone_data,
    dict_to_lynks_data,
    BASE_DIR,
)
from src.pptx_generator import generate_pptx
from src.ocr_extractor import extract_crm_data, format_crm_data_as_yaml

HISTORY_DIR = BASE_DIR / "history"


def save_to_history(brand: str, client_data: dict, output_path: Path):
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    brand_dir = HISTORY_DIR / brand
    brand_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    client_name = _get_client_name(client_data, brand)
    history_file = brand_dir / f"{client_name}_{timestamp}.yaml"
    history_data = {
        "timestamp": timestamp,
        "brand": brand,
        "output_file": str(output_path),
        "data": client_data,
    }
    with open(history_file, "w", encoding="utf-8") as f:
        yaml.dump(history_data, f, default_flow_style=False, allow_unicode=True, sort_keys=False)
    print(f"[HIST] Datos guardados en historial: {history_file}")


def load_previous_history(brand: str, client_name: str) -> dict | None:
    brand_dir = HISTORY_DIR / brand
    if not brand_dir.exists():
        return None
    history_files = sorted(brand_dir.glob(f"{client_name}_*.yaml"), reverse=True)
    for hf in history_files:
        if hf.is_file():
            with open(hf, "r", encoding="utf-8") as f:
                return yaml.safe_load(f)
    return None


def compute_diff(current: dict, previous: dict) -> list[dict]:
    diffs = []
    _diff_dicts(current, previous, "", diffs)
    return diffs


def _diff_dicts(current, previous, path, diffs):
    if isinstance(current, list) and isinstance(previous, list):
        if current != previous:
            for i, (c_item, p_item) in enumerate(zip(current, previous)):
                if isinstance(c_item, dict) and isinstance(p_item, dict):
                    _diff_dicts(c_item, p_item, f"{path}[{i}]", diffs)
                elif c_item != p_item:
                    diffs.append({"field": f"{path}[{i}]", "old": p_item, "new": c_item, "change": "changed"})
            if len(current) > len(previous):
                for i in range(len(previous), len(current)):
                    diffs.append({"field": f"{path}[{i}]", "old": "(nuevo)", "new": current[i], "change": "added"})
            elif len(current) < len(previous):
                for i in range(len(current), len(previous)):
                    diffs.append({"field": f"{path}[{i}]", "old": previous[i], "new": "(eliminado)", "change": "removed"})
        return

    all_keys = set(list(current.keys()) if isinstance(current, dict) else [] + list(previous.keys()) if isinstance(previous, dict) else [])
    if not isinstance(current, dict) or not isinstance(previous, dict):
        return

    for key in all_keys:
        current_path = f"{path}.{key}" if path else key
        curr_val = current.get(key)
        prev_val = previous.get(key)

        if curr_val is None and prev_val is not None:
            diffs.append({"field": current_path, "old": prev_val, "new": "(eliminado)", "change": "removed"})
        elif curr_val is not None and prev_val is None:
            diffs.append({"field": current_path, "old": "(nuevo)", "new": curr_val, "change": "added"})
        elif isinstance(curr_val, dict) and isinstance(prev_val, dict):
            _diff_dicts(curr_val, prev_val, current_path, diffs)
        elif isinstance(curr_val, list) and isinstance(prev_val, list):
            _diff_dicts(curr_val, prev_val, current_path, diffs)
        elif curr_val != prev_val:
            diffs.append({"field": current_path, "old": prev_val, "new": curr_val, "change": "changed"})


def format_diff(diffs: list[dict]) -> str:
    if not diffs:
        return "  Sin cambios respecto a la version anterior."
    lines = []
    price_fields = {"price_monthly", "price_before", "total_monthly", "apoyo_economico", "total_price"}
    for d in diffs:
        field = d["field"]
        icon = "€" if any(p in field for p in price_fields) else ("+" if d["change"] == "added" else "-")
        if d["change"] == "changed":
            lines.append(f"  {icon} {field}: {d['old']} → {d['new']}")
        elif d["change"] == "added":
            lines.append(f"  + {field}: {d['new']}")
        elif d["change"] == "removed":
            lines.append(f"  - {field}: {d['old']}")
    return "\n".join(lines)


def confirm_data(client_data: dict, brand: str) -> bool:
    print("\n" + "=" * 60)
    print(f"  RESUMEN DE DATOS — {brand.upper()}")
    print("=" * 60)

    client = client_data.get("client", {})
    proposal = client_data.get("proposal", {})
    products = client_data.get("products", [])
    economic = client_data.get("economic", {})

    if brand == "vodafone":
        print(f"\n  Cliente: {client.get('name_line1', '')} {client.get('name_line2', '')}")
        print(f"  NIF: {client.get('nif', '')}")
        print(f"  Lineas: {client.get('lines', '')}")
        print(f"  Tipo: {client.get('type', '')}")
        print(f"  Fecha: {proposal.get('date', '')}")
        print(f"  Permanencia: {proposal.get('duration_months', '')}")
        if products:
            print(f"\n  PRODUCTOS:")
            for i, p in enumerate(products[:3], 1):
                print(f"    {i}. {p.get('name_line1', '')} {p.get('name_line2', '')}")
                print(f"       Precio: {p.get('price_monthly', '')}")
                print(f"       Antes: {p.get('price_before', '')}")
                print(f"       Descuento: {p.get('discount_text', '')}")
        if economic:
            print(f"\n  ECONOMICO:")
            print(f"    Total: {economic.get('total_monthly', '')}")
            print(f"    Apoyo: {economic.get('apoyo_economico', '')}")
    else:
        print(f"\n  Cliente: {client.get('name_short', '')} {client.get('name_suffix', '')}")
        print(f"  NIF: {client.get('nif', '')}")
        print(f"  Sedes: {client.get('sedes', '')}")
        print(f"  Extensiones: {proposal.get('extensiones', '')}")
        print(f"  Fecha: {proposal.get('date', '')}")
        options = client_data.get("options", [])
        if options:
            print(f"\n  OPCIONES:")
            for opt in options:
                print(f"    {opt.get('title', '')}: {opt.get('total_price', '')}")

    print("\n" + "=" * 60)

    previous = load_previous_history(brand, _get_client_name(client_data, brand))
    if previous and previous.get("data"):
        diffs = compute_diff(client_data, previous.get("data", {}))
        if diffs:
            print("  CAMBIOS RESPECTO A VERSION ANTERIOR:")
            print(format_diff(diffs))
        else:
            print("  Sin cambios respecto a la version anterior.")
        print()

    try:
        response = input("\n  ¿Generar presentacion? [S/n]: ").strip().lower()
        return response in ("s", "si", "sí", "y", "yes", "")
    except (EOFError, KeyboardInterrupt):
        print("\n  Cancelado.")
        return False


def _get_client_name(data: dict, brand: str) -> str:
    if brand == "vodafone":
        c = data.get("client", {})
        return f"{c.get('name_line1', 'Cliente')}_{c.get('name_line2', '')}".strip("_")
    c = data.get("client", {})
    return f"{c.get('name_short', 'Cliente')}_{c.get('name_suffix', '')}".strip("_")


def main():
    parser = argparse.ArgumentParser(
        description="Generador automatico de presentaciones de precios"
    )
    parser.add_argument("--brand", choices=["vodafone", "lynks-tic"], required=True, help="Marca")
    parser.add_argument("--data", help="Ruta al archivo YAML con datos del cliente")
    parser.add_argument("--screenshot", help="Ruta a la captura del CRM para OCR")
    parser.add_argument("--override", help="Ruta al archivo YAML de override manual")
    parser.add_argument("--include-roaming", action="store_true", help="Incluir slides de roaming (Vodafone)")
    parser.add_argument("--output", help="Nombre del archivo de salida")
    parser.add_argument("--dry-run", action="store_true", help="Solo mostrar datos sin generar")
    parser.add_argument("--save-yaml", help="Guardar datos extraidos como YAML")
    parser.add_argument("--interactive", "-i", action="store_true", help="Modo interactivo: confirmar datos antes de generar")
    parser.add_argument("--diff", action="store_true", help="Mostrar diff con version anterior")
    parser.add_argument("--no-history", action="store_true", help="No guardar en historial")

    args = parser.parse_args()

    if not args.data and not args.screenshot:
        parser.error("Debes proporcionar --data o --screenshot (o ambos)")

    client_data = {}

    if args.screenshot:
        print(f"[OCR] Extrayendo datos de: {args.screenshot}")
        ocr_result = extract_crm_data(args.screenshot)
        print(f"[OCR] Encontrados {len(ocr_result['prices'])} precios")
        print(f"[OCR] Encontrados {ocr_result['lines']['lines']} lineas")

        if args.save_yaml:
            yaml_path = Path(args.save_yaml)
            yaml_path.parent.mkdir(parents=True, exist_ok=True)
            yaml_content = format_crm_data_as_yaml(ocr_result)
            yaml_path.write_text(yaml_content, encoding="utf-8")
            print(f"[OCR] Datos guardados en: {yaml_path}")

        client_data = _ocr_to_client_data(ocr_result, args.brand)

    if args.data:
        print(f"[DATA] Cargando datos de: {args.data}")
        file_data = load_client_data(args.data)
        client_data.update(file_data)

    if args.override:
        print(f"[OVERRIDE] Aplicando override: {args.override}")
        client_data = merge_overrides(client_data, args.override)

    if args.include_roaming:
        proposal = client_data.get("proposal", client_data)
        if isinstance(proposal, dict):
            proposal["include_roaming"] = True
            client_data["proposal"] = proposal
        else:
            client_data["include_roaming"] = True

    if args.diff or args.interactive:
        previous = load_previous_history(args.brand, _get_client_name(client_data, args.brand))
        if previous and previous.get("data"):
            diffs = compute_diff(client_data, previous.get("data", {}))
            if diffs:
                print("\n[DIFF] Cambios respecto a version anterior:")
                print(format_diff(diffs))
            else:
                print("\n[DIFF] Sin cambios respecto a version anterior.")

    if args.interactive:
        if not confirm_data(client_data, args.brand):
            print("[CANCELADO] Generacion cancelada por el usuario.")
            return

    if args.dry_run:
        print("\n[DRY-RUN] Datos que se usarian:")
        print(json.dumps(client_data, indent=2, ensure_ascii=False, default=str))
        return

    print(f"[GEN] Generando presentacion {args.brand}...")
    output_path = generate_pptx(
        brand=args.brand,
        client_data=client_data,
        output_name=args.output,
    )
    print(f"[OK] Presentacion generada: {output_path}")

    if not args.no_history:
        save_to_history(args.brand, client_data, output_path)


def _ocr_to_client_data(ocr_data: dict, brand: str) -> dict:
    result = {}

    if ocr_data["client"].get("nif"):
        result.setdefault("client", {})["nif"] = ocr_data["client"]["nif"]

    if ocr_data["lines"].get("lines"):
        result.setdefault("client", {})["lines"] = ocr_data["lines"]["lines"]

    if ocr_data["prices"]:
        result.setdefault("products", [])
        for price_info in ocr_data["prices"][:5]:
            result["products"].append({
                "price_monthly": price_info["original"]
            })

    return result


if __name__ == "__main__":
    main()