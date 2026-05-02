import yaml
from pathlib import Path
from src.models import VodafoneClientData, LynksTICClientData, ProductInfo, TerminalInfo, EconomicInfo, LynksTICOption, FieldMapping


BASE_DIR = Path(__file__).resolve().parent.parent


def load_settings() -> dict:
    settings_path = BASE_DIR / "config" / "settings.yaml"
    with open(settings_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_brand_config(brand: str) -> dict:
    settings = load_settings()
    config_path = BASE_DIR / settings["brand_configs"][brand]
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_field_mappings(brand: str) -> dict[str, list[FieldMapping]]:
    config = load_brand_config(brand)
    mappings = {}
    for slide_key, fields in config.get("fields", {}).items():
        slide_mappings = []
        for field_name, field_def in fields.items():
            slide_mappings.append(FieldMapping(
                path=field_def["path"],
                field=field_name,
                find=field_def["find"],
                type=field_def.get("type", "text")
            ))
        mappings[slide_key] = slide_mappings
    return mappings


def load_client_data(data_path: str) -> dict:
    data_file = Path(data_path)
    if not data_file.exists():
        raise FileNotFoundError(f"Client data file not found: {data_path}")
    with open(data_file, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def merge_overrides(base_data: dict, overrides_path: str) -> dict:
    override_file = Path(overrides_path)
    if not override_file.exists():
        return base_data
    with open(override_file, "r", encoding="utf-8") as f:
        overrides = yaml.safe_load(f)
    if not overrides:
        return base_data
    return _deep_merge(base_data, overrides)


def _deep_merge(base: dict, override: dict) -> dict:
    result = base.copy()
    for key, value in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def dict_to_vodafone_data(data: dict) -> VodafoneClientData:
    products = []
    for p in data.get("products", []):
        products.append(ProductInfo(
            name_line1=p.get("name_line1", ""),
            name_line2=p.get("name_line2", ""),
            price_monthly=p.get("price_monthly", ""),
            price_before=p.get("price_before", ""),
            discount_text=p.get("discount_text", ""),
            features=p.get("features", [])
        ))

    terminals = []
    for t in data.get("terminals", []):
        terminals.append(TerminalInfo(
            model=t.get("model", ""),
            price=t.get("price", "0,00 €"),
            specs=t.get("specs", {})
        ))

    economic = EconomicInfo(**data.get("economic", {}))

    client = data.get("client", {})
    proposal = data.get("proposal", {})

    return VodafoneClientData(
        client_name_line1=client.get("name_line1", ""),
        client_name_line2=client.get("name_line2", ""),
        nif=client.get("nif", ""),
        lines=client.get("lines", 0),
        client_type=client.get("type", "Cliente Nuevo"),
        permanencia=proposal.get("duration_months", "36 meses"),
        network_type=proposal.get("network_type", "Red 5G Vodafone"),
        feature_bullets=client.get("feature_bullets", ""),
        date=proposal.get("date", "Abril 2026"),
        products=products,
        terminals=terminals,
        economic=economic,
        include_holidaysim=proposal.get("include_holidaysim", True),
        include_lynks_tic=proposal.get("include_lynks_tic", True),
        include_roaming=proposal.get("include_roaming", False),
    )


def dict_to_lynks_data(data: dict) -> LynksTICClientData:
    client = data.get("client", {})
    proposal = data.get("proposal", {})
    options = []
    for opt in data.get("options", []):
        options.append(LynksTICOption(**opt))

    return LynksTICClientData(
        client_name_short=client.get("name_short", ""),
        client_name_suffix=client.get("name_suffix", ""),
        nif=client.get("nif", ""),
        sedes=client.get("sedes", 0),
        service_type=client.get("service_type", "Conectividad FTTO + Voz IP"),
        date=proposal.get("date", "Abril 2026"),
        extensiones=proposal.get("extensiones", 0),
        canales=proposal.get("canales", 0),
        ddi_unit_price=proposal.get("ddi_unit_price", "1,50€/mes"),
        terminal_model=proposal.get("terminal_model", "Yealink W73P"),
        terminal_unit_price=proposal.get("terminal_unit_price", "8,50€/ud/mes"),
        options=options,
    )