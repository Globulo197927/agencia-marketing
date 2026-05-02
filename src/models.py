from dataclasses import dataclass, field
from typing import Optional


@dataclass
class ProductInfo:
    name_line1: str = ""
    name_line2: str = ""
    price_monthly: str = ""
    price_before: str = ""
    discount_text: str = ""
    features: list[str] = field(default_factory=list)


@dataclass
class TerminalInfo:
    model: str = ""
    price: str = "0,00 €"
    specs: dict[str, str] = field(default_factory=dict)


@dataclass
class EconomicInfo:
    total_monthly: str = ""
    subsidy_vodafone: str = ""
    subsidy_plan_negocio: str = ""
    apoyo_economico: str = ""
    summary_line: str = ""
    subsidy_detail_line: str = ""


@dataclass
class VodafoneClientData:
    client_name_line1: str = ""
    client_name_line2: str = ""
    nif: str = ""
    lines: int = 0
    client_type: str = "Cliente Nuevo"
    permanencia: str = "36 meses"
    network_type: str = "Red 5G Vodafone"
    feature_bullets: str = ""
    date: str = ""

    products: list[ProductInfo] = field(default_factory=list)
    terminals: list[TerminalInfo] = field(default_factory=list)
    economic: EconomicInfo = field(default_factory=EconomicInfo)

    include_holidaysim: bool = True
    include_lynks_tic: bool = True
    include_roaming: bool = False


@dataclass
class LynksTICOption:
    name: str = ""
    speed: str = ""
    speed_download: str = ""
    speed_upload: str = ""
    total_price: str = ""
    ftto_units: int = 0
    ftto_unit_price: str = ""
    ftto_total: str = ""


@dataclass
class LynksTICClientData:
    client_name_short: str = ""
    client_name_suffix: str = ""
    nif: str = ""
    sedes: int = 0
    service_type: str = "Conectividad FTTO + Voz IP"
    date: str = ""

    extensiones: int = 0
    canales: int = 0
    ddi_unit_price: str = "1,50€/mes"
    terminal_model: str = "Yealink W73P"
    terminal_unit_price: str = "8,50€/ud/mes"

    options: list[LynksTICOption] = field(default_factory=list)

    option_a_title: str = ""
    option_b_title: str = ""
    savings_line: str = ""


@dataclass
class FieldMapping:
    path: str
    field: str
    find: str
    type: str = "text"

    def get_replace_value(self, data) -> str:
        if hasattr(data, self.field):
            return str(getattr(data, self.field))
        keys = self.field.split(".")
        obj = data
        for key in keys:
            if isinstance(obj, dict):
                obj = obj.get(key, "")
            elif hasattr(obj, key):
                obj = getattr(obj, key)
            else:
                return ""
        return str(obj) if obj is not None else ""