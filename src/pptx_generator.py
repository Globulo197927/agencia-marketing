import subprocess
import json
import shutil
from pathlib import Path
from datetime import datetime
from pptx import Presentation
from pptx.util import Cm, Emu

from src.config_loader import load_brand_config, BASE_DIR


def _run_officecli(args: list[str]) -> str:
    cmd = ["officecli"] + args
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if result.returncode != 0:
        raise RuntimeError(f"officecli error: {result.stderr}")
    return result.stdout


def _path(p: Path) -> str:
    return str(p.resolve())


def generate_pptx(brand: str, client_data: dict, output_name: str = None) -> Path:
    config = load_brand_config(brand)
    template_path = BASE_DIR / config["brand"]["template"]
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    brand_dir = "vodafone" if brand == "vodafone" else "lynks-tic"
    output_dir = BASE_DIR / "output" / brand_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    if output_name is None:
        client_name = _get_client_name(client_data, brand)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_name = f"{brand}_{client_name}_{timestamp}.pptx"

    if not output_name.endswith(".pptx"):
        output_name += ".pptx"

    output_path = output_dir / output_name
    shutil.copy2(template_path, output_path)

    if brand == "vodafone":
        commands = _build_vodafone_commands(client_data, str(output_path.resolve()))
    else:
        commands = _build_lynks_commands(client_data, str(output_path.resolve()))

    if commands:
        batch_json = json.dumps(commands, ensure_ascii=False)
        _run_officecli(["batch", str(output_path.resolve()), "--commands", batch_json, "--json"])

    _handle_optional_slides(output_path, client_data, brand)
    _add_terminal_images(output_path, client_data, brand)

    return output_path


def _set(path: str, find_text: str, replace_text: str) -> dict:
    if find_text == replace_text:
        return None
    return {"command": "set", "path": path, "props": {"find": find_text, "replace": replace_text}}


def _get_client_name(data: dict, brand: str) -> str:
    if brand == "vodafone":
        c = data.get("client", {})
        return f"{c.get('name_line1', 'Cliente')}_{c.get('name_line2', '')}".strip("_")
    c = data.get("client", {})
    return f"{c.get('name_short', 'Cliente')}_{c.get('name_suffix', '')}".strip("_")


def _build_vodafone_commands(data: dict, filepath: str) -> list[dict]:
    c = data.get("client", {})
    p = data.get("proposal", {})
    prods = data.get("products", [])
    eco = data.get("economic", {})
    terms = data.get("terminals", [])

    full_name = f"{c.get('name_line1', '')} {c.get('name_line2', '')}".strip()
    lines = c.get("lines", 25)
    ctype = c.get("type", "Cliente Nuevo")
    date = p.get("date", "Marzo 2026")
    net = p.get("network_type", "Red 5G Vodafone")

    cmds = []

    # Slide 1
    cmds.append(_set("/slide[1]/shape[@id=6]", "TECNOLOGÍA", c.get("name_line1", "TECNOLOGÍA")))
    cmds.append(_set("/slide[1]/shape[@id=7]", "TÉRMICA SA", c.get("name_line2", "TÉRMICA SA")))
    cmds.append(_set("/slide[1]/shape[@id=9]", "NIF A28592772", f"NIF {c.get('nif', 'A28592772')}"))
    cmds.append(_set("/slide[1]/shape[@id=10]", "● Red 5G Vodafone", f"● {net}"))
    cmds.append(_set("/slide[1]/shape[@id=11]", "Abril 2026", date))
    cmds.append(_set("/slide[1]/shape[@id=13]", "Marzo 2026", date))

    bullets = c.get("feature_bullets", "")
    if bullets:
        original_bullets = "▸  25 líneas móviles · Plan RED Infinity PRO\n▸  Fibra 1Gbps + Centralita Plus incluidas\n▸  Terminales a precio 0€ disponibles"
        cmds.append(_set("/slide[1]/shape[@id=12]", original_bullets, bullets))
    elif prods:
        p1 = prods[0]
        new_bullets = f"▸  {lines} líneas móviles · {p1.get('name_line1', '')} {p1.get('name_line2', '')}\n▸  Fibra 1Gbps + Centralita Plus incluidas\n▸  Terminales a precio 0€ disponibles"
        old_bullets = "▸  25 líneas móviles · Plan RED Infinity PRO\n▸  Fibra 1Gbps + Centralita Plus incluidas\n▸  Terminales a precio 0€ disponibles"
        cmds.append(_set("/slide[1]/shape[@id=12]", old_bullets, new_bullets))

    # Slide 2
    cmds.append(_set("/slide[2]/shape[@id=28]", "TECNOLOGÍA TÉRMICA SA", full_name))

    # Slide 5
    summary = f"Servicios base para {full_name} · {lines} líneas · {ctype} · Compromiso {p.get('duration_months', '36 meses')}"
    cmds.append(_set("/slide[5]/shape[@id=4]", "Servicios base para TECNOLOGÍA TÉRMICA SA · 25 líneas · Cliente Nuevo · Compromiso 36 meses", summary))

    if len(prods) >= 1:
        p1 = prods[0]
        cmds.append(_set("/slide[5]/shape[@id=8]", f"BUNDLE RED\nINFINITY PRO", f"{p1.get('name_line1', 'BUNDLE RED')}\n{p1.get('name_line2', 'INFINITY PRO')}"))
        cmds.append(_set("/slide[5]/shape[@id=10]", "73,44€/mes", p1.get("price_monthly", "73,44€/mes")))
        cmds.append(_set("/slide[5]/shape[@id=11]", "antes: 112,98€/mes", f"antes: {p1.get('price_before', '112,98€/mes')}"))
        cmds.append(_set("/slide[5]/shape[@id=13]", "35% de descuento", p1.get("discount_text", "35% de descuento")))
        feats = p1.get("features", [])
        default_feats_1 = ["Datos ilimitados Full Speed", "5G en toda la red Vodafone", "Roaming incluido en UE"]
        for i in range(3):
            shape_ids = [15, 17, 19]
            if i < len(feats):
                cmds.append(_set(f"/slide[5]/shape[@id={shape_ids[i]}]", default_feats_1[i], feats[i]))

    if len(prods) >= 2:
        p2 = prods[1]
        cmds.append(_set("/slide[5]/shape[@id=23]", f"FIBRA\n1 GBPS", f"{p2.get('name_line1', 'FIBRA')}\n{p2.get('name_line2', '1 GBPS')}"))
        cmds.append(_set("/slide[5]/shape[@id=25]", "0€/mes", p2.get("price_monthly", "0€/mes")))
        cmds.append(_set("/slide[5]/shape[@id=26]", "antes: 34,85€/mes", f"antes: {p2.get('price_before', '34,85€/mes')}"))
        cmds.append(_set("/slide[5]/shape[@id=28]", "100% de descuento", p2.get("discount_text", "100% de descuento")))
        feats2 = p2.get("features", [])
        default_feats_2 = ["Conexión simétrica 1 Gbps", "Incluida sin coste adicional", "Instalación y mantenimiento"]
        shape_ids_2 = [30, 32, 34]
        for i in range(3):
            if i < len(feats2):
                cmds.append(_set(f"/slide[5]/shape[@id={shape_ids_2[i]}]", default_feats_2[i], feats2[i]))

    if len(prods) >= 3:
        p3 = prods[2]
        cmds.append(_set("/slide[5]/shape[@id=38]", f"CENTRALITA\nPLUS", f"{p3.get('name_line1', 'CENTRALITA')}\n{p3.get('name_line2', 'PLUS')}"))
        cmds.append(_set("/slide[5]/shape[@id=40]", "0€/mes", p3.get("price_monthly", "0€/mes")))
        cmds.append(_set("/slide[5]/shape[@id=41]", "antes: 30,75€/mes", f"antes: {p3.get('price_before', '30,75€/mes')}"))
        cmds.append(_set("/slide[5]/shape[@id=43]", "100% de descuento", p3.get("discount_text", "100% de descuento")))
        feats3 = p3.get("features", [])
        default_feats_3 = ["Gestión avanzada de llamadas", "Incluida sin coste adicional", "Multi-sede y extensiones"]
        shape_ids_3 = [45, 47, 49]
        for i in range(3):
            if i < len(feats3):
                cmds.append(_set(f"/slide[5]/shape[@id={shape_ids_3[i]}]", default_feats_3[i], feats3[i]))

    cmds.append(_set("/slide[5]/shape[@id=51]", "25 líneas móviles", f"{lines} líneas móviles"))
    cmds.append(_set("/slide[5]/shape[@id=53]", "Red 5G Vodafone", net))
    cmds.append(_set("/slide[5]/shape[@id=55]", "Cliente Nuevo", ctype))
    cmds.append(_set("/slide[5]/shape[@id=61]", "1.800€", eco.get("apoyo_economico", "1.800€")))

    # Slide 6
    if eco:
        cmds.append(_set("/slide[6]/shape[@id=5]", "305,79 €/mes", eco.get("total_monthly", "305,79 €/mes")))
        cmds.append(_set("/slide[6]/shape[@id=6]", "25 líneas · Terminales a 0€ · Todos con 35% dto · Compromiso 36 meses", eco.get("summary_line", "")))
        cmds.append(_set("/slide[6]/shape[@id=7]", "Sub. Vodafone: -160,82€  ·  Sub. Plan Negocio: -956,18€  ·  Apoyo Económico: 1.800€", eco.get("subsidy_detail_line", "")))

    # Slide 6 — Table cells (distribution table)
    table_data = data.get("distribution_table", [])
    if table_data:
        default_table = [
            {"terminal": "Adoc NEO 8000 4G Deals", "plan": "RED Infinity PRO Ilim. Bundle", "precio_cesion": "187,00€", "sin_dto": "0,00€", "dto": "35%", "con_dto": "0,00€", "uds": "1", "total": "0,00€"},
            {"terminal": "Adoc NEO 3850W 4G Deals", "plan": "RED Infinity PRO Ilim. Full Speed", "precio_cesion": "365,00€", "sin_dto": "17,95€", "dto": "35%", "con_dto": "11,67€", "uds": "5", "total": "58,34€"},
            {"terminal": "Adoc NEO 3850W 4G Deals", "plan": "RED Infinity PRO Ilim. Full Speed", "precio_cesion": "365,00€", "sin_dto": "17,95€", "dto": "35%", "con_dto": "11,67€", "uds": "5", "total": "58,34€"},
            {"terminal": "—", "plan": "RED Infinity PRO 5Gb", "precio_cesion": "0,00€", "sin_dto": "13,73€", "dto": "35%", "con_dto": "8,92€", "uds": "10", "total": "89,24€"},
            {"terminal": "—", "plan": "RED Infinity PRO 5Gb", "precio_cesion": "0,00€", "sin_dto": "13,73€", "dto": "35%", "con_dto": "8,92€", "uds": "1", "total": "8,92€"},
            {"terminal": "Cocomm DT250 4G Solo Voz", "plan": "RED Infinity PRO Solo Voz", "precio_cesion": "198,00€", "sin_dto": "8,98€", "dto": "35%", "con_dto": "5,84€", "uds": "3", "total": "17,51€"},
        ]
        for row_idx, row_data in enumerate(table_data):
            if row_idx >= 6:
                break
            tr = row_idx + 2  # rows 2-7
            default_row = default_table[row_idx] if row_idx < len(default_table) else {}
            for col_idx, key in enumerate(["terminal", "plan", "precio_cesion", "sin_dto", "dto", "con_dto", "uds", "total"]):
                tc = col_idx + 1
                default_value = default_row.get(key, "")
                new_value = row_data.get(key, default_value)
                if default_value and new_value:
                    cmds.append(_set(f"/slide[6]/table[@id=8]/tr[{tr}]/tc[{tc}]", default_value, new_value))

    # Total row of distribution table
    line_data = data.get("distribution_total", {})
    if line_data:
        uds_total = line_data.get("uds", str(lines) + " uds")
        total_price = line_data.get("total", eco.get("total_monthly", "305,79€") if eco else "305,79€")
        cmds.append(_set("/slide[6]/table[@id=8]/tr[8]/tc[7]", "25 uds", uds_total))
        cmds.append(_set("/slide[6]/table[@id=8]/tr[8]/tc[8]", "305,79€", total_price))

    # Slide 7
    if terms:
        t1 = terms[0]
        cmds.append(_set("/slide[7]/shape[@id=7]", "Adoc NEO 8000 4G", t1.get("model", "Adoc NEO 8000 4G")))
        cmds.append(_set("/slide[7]/shape[@id=9]", "PRECIO CESIÓN: 0,00 €", f"PRECIO CESIÓN: {t1.get('price', '0,00 €')}"))
    if len(terms) >= 2:
        t2 = terms[1]
        cmds.append(_set("/slide[7]/shape[@id=37]", "Adoc NEO 3850W 4G", t2.get("model", "Adoc NEO 3850W 4G")))
        cmds.append(_set("/slide[7]/shape[@id=39]", "PRECIO CESIÓN: 0,00 €", f"PRECIO CESIÓN: {t2.get('price', '0,00 €')}"))
    if len(terms) >= 3:
        t3 = terms[2]
        cmds.append(_set("/slide[7]/shape[@id=67]", "Cocomm DT250 4G Solo Voz", t3.get("model", "Cocomm DT250 4G Solo Voz")))
        cmds.append(_set("/slide[7]/shape[@id=69]", "PRECIO CESIÓN: 0,00 €", f"PRECIO CESIÓN: {t3.get('price', '0,00 €')}"))

    # Slide 8
    cmds.append(_set("/slide[8]/shape[@id=4]", "TECNOLOGÍA TÉRMICA SA", full_name))

    return [c for c in cmds if c is not None]


def _build_lynks_commands(data: dict, filepath: str) -> list[dict]:
    c = data.get("client", {})
    p = data.get("proposal", {})
    opts = data.get("options", [])

    short = c.get("name_short", "DIAVERUM")
    suffix = c.get("name_suffix", "SERVICIOS RENALES SL")
    full = f"{short} {suffix}".strip()
    sedes = c.get("sedes", 55)
    nif = c.get("nif", "B84991736")
    date = p.get("date", "Marzo 2026")
    ext = p.get("extensiones", 165)

    cmds = []

    # Slide 1
    cmds.append(_set("/slide[1]/shape[@id=6]", "DIAVERUM", short))
    cmds.append(_set("/slide[1]/shape[@id=7]", "SERVICIOS RENALES SL", suffix))
    cmds.append(_set("/slide[1]/shape[@id=9]", "B84991736  ·  55 sedes  ·  Conectividad FTTO + Voz IP  ·  Marzo 2026",
                      f"{nif}  ·  {sedes} sedes  ·  {c.get('service_type', 'Conectividad FTTO + Voz IP')}  ·  {date}"))

    # Slides 2, 4
    cmds.append(_set("/slide[2]/shape[@id=29]", "Diaverum", short))
    cmds.append(_set("/slide[4]/shape[@id=5]", "Diaverum Servicios Renales SL", full))

    # Slide 12
    cmds.append(_set("/slide[12]/shape[@id=5]",
                      "165 unidades · 8,50€/ud/mes · Tarifa plana ilimitada a fijos y móviles nacionales",
                      f"{ext} unidades · {p.get('terminal_unit_price', '8,50€/ud/mes')} · Tarifa plana ilimitada a fijos y móviles nacionales"))

    # Slide 13
    cmds.append(_set("/slide[13]/shape[@id=5]", "las 55 sedes de Diaverum Servicios Renales SL", f"las {sedes} sedes de {full}"))
    cmds.append(_set("/slide[13]/shape[@id=9]", "165", str(ext)))
    cmds.append(_set("/slide[13]/shape[@id=17]", "165", str(ext)))
    cmds.append(_set("/slide[13]/shape[@id=25]", "55", str(sedes)))
    cmds.append(_set("/slide[13]/shape[@id=41]", "165", str(ext)))

    # Slide 14 - Option A
    if len(opts) >= 1:
        a = opts[0]
        cmds.append(_set("/slide[14]/shape[@id=4]", "OPCIÓN A — FTTO 1 GBPS · LA MÁS POTENTE", a.get("title", "OPCIÓN A — FTTO 1 GBPS · LA MÁS POTENTE")))
        cmds.append(_set("/slide[14]/shape[@id=6]", "3.795 €/mes", a.get("total_price", "3.795 €/mes")))
        cmds.append(_set("/slide[14]/shape[@id=7]", "55 sedes · 1Gbps/1Gbps simétrico · Compromiso anual · SLA 99,9%",
                          f"{sedes} sedes · {a.get('speed', '1Gbps/1Gbps simétrico')} · Compromiso anual · SLA 99,9%"))
        cmds.append(_set("/slide[14]/shape[@id=8]", "Descarga 1.000 Mbps · Subida 1.000 Mbps · Latencia <10ms",
                          f"Descarga {a.get('speed_download', '1.000 Mbps')} · Subida {a.get('speed_upload', '1.000 Mbps')} · Latencia <10ms"))

    # Slide 15 - Option B
    if len(opts) >= 2:
        b = opts[1]
        cmds.append(_set("/slide[15]/shape[@id=4]", "OPCIÓN B — FTTO 600 MBPS · LA MÁS EQUILIBRADA", b.get("title", "OPCIÓN B — FTTO 600 MBPS · LA MÁS EQUILIBRADA")))
        cmds.append(_set("/slide[15]/shape[@id=6]", "3.575 €/mes", b.get("total_price", "3.575 €/mes")))
        cmds.append(_set("/slide[15]/shape[@id=7]", "55 sedes · 600Mbps/600Mbps simétrico · Compromiso anual · SLA 99,9%",
                          f"{sedes} sedes · {b.get('speed', '600Mbps/600Mbps simétrico')} · Compromiso anual · SLA 99,9%"))

    # Slide 16 - Comparison
    if len(opts) >= 1:
        a = opts[0]
        cmds.append(_set("/slide[16]/shape[@id=9]", "FTTO 1 Gbps simétrico", a.get("speed_label", "FTTO 1 Gbps simétrico")))
        cmds.append(_set("/slide[16]/shape[@id=11]", "3.795 €/mes", a.get("total_price", "3.795 €/mes")))
        cmds.append(_set("/slide[16]/shape[@id=13]", "55 × FTTO 1 Gbps — 2.310€", f"{sedes} × FTTO 1 Gbps — {a.get('ftto_total', '2.310€')}"))

    if len(opts) >= 2:
        b = opts[1]
        cmds.append(_set("/slide[16]/shape[@id=26]", "FTTO 600 Mbps simétrico", b.get("speed_label", "FTTO 600 Mbps simétrico")))
        cmds.append(_set("/slide[16]/shape[@id=28]", "3.575 €/mes", b.get("total_price", "3.575 €/mes")))
        cmds.append(_set("/slide[16]/shape[@id=30]", "55 × FTTO 600 Mbps — 2.090€", f"{sedes} × FTTO 600 Mbps — {b.get('ftto_total', '2.090€')}"))

    cmds.append(_set("/slide[16]/shape[@id=15]", "55 × DDI — 82,50€", f"{sedes} × DDI — {sedes * 1.5:.2f}€"))
    cmds.append(_set("/slide[16]/shape[@id=32]", "55 × DDI — 82,50€", f"{sedes} × DDI — {sedes * 1.5:.2f}€"))
    cmds.append(_set("/slide[16]/shape[@id=17]", "165 × Yealink W37P — 1.402,50€", f"{ext} × Yealink W37P — {ext * 8.5:.2f}€"))
    cmds.append(_set("/slide[16]/shape[@id=34]", "165 × Yealink W37P — 1.402,50€", f"{ext} × Yealink W37P — {ext * 8.5:.2f}€"))

    savings = data.get("savings_line", "↓ Ahorro de 2.640€/año respecto a Opción A")
    cmds.append(_set("/slide[16]/shape[@id=37]", "↓ Ahorro de 2.640€/año respecto a Opción A", savings))

    # Slide 17
    cmds.append(_set("/slide[17]/shape[@id=5]", "DIAVERUM SERVICIOS RENALES SL", full))

    return [c for c in cmds if c is not None]


TERMINAL_IMAGES = {
    "vodafone": {
        "slide": 7,
        "terminals": [
            {"key": "adoc_neo_8000", "image": "assets/adoc_neo_8000.png", "x": "5.2cm", "y": "2.7cm", "w": "2.5cm", "h": "2.5cm"},
            {"key": "adoc_neo_3850w", "image": "assets/adoc_neo_3850w.png", "x": "13.6cm", "y": "2.7cm", "w": "2.5cm", "h": "2.5cm"},
            {"key": "cocomm_dt250", "image": "assets/cocomm_dt250.png", "x": "21.8cm", "y": "2.7cm", "w": "2.5cm", "h": "2.5cm"},
        ],
    },
    "lynks-tic": {
        "slide": 12,
        "terminals": [
            {"key": "yealink_w73p", "image": "assets/yealink_w73p.png", "picture_id": 8, "x": "1.83cm", "y": "3.05cm", "w": "6.1cm", "h": "10.16cm"},
        ],
    },
}


def _parse_cm(value: str) -> int:
    return Cm(float(value.replace("cm", "")))


def _add_terminal_images(output_path: Path, client_data: dict, brand: str):
    config = TERMINAL_IMAGES.get(brand)
    if not config:
        return

    img_path = output_path.resolve()
    prs = Presentation(str(img_path))
    slide_idx = config["slide"] - 1
    slide = prs.slides[slide_idx]

    if brand == "vodafone":
        for i, term_conf in enumerate(config["terminals"]):
            ipath = BASE_DIR / term_conf["image"]
            if not ipath.exists():
                continue
            slide.shapes.add_picture(
                str(ipath),
                _parse_cm(term_conf["x"]),
                _parse_cm(term_conf["y"]),
                _parse_cm(term_conf["w"]),
                _parse_cm(term_conf["h"]),
            )
    elif brand == "lynks-tic":
        for term_conf in config["terminals"]:
            ipath = BASE_DIR / term_conf["image"]
            if not ipath.exists():
                continue
            pid = term_conf.get("picture_id")
            if pid:
                for shape in list(slide.shapes):
                    try:
                        if hasattr(shape, 'image') and shape.name == "Image 0":
                            sp = shape._element
                            sp.getparent().remove(sp)
                            break
                    except Exception:
                        pass
            slide.shapes.add_picture(
                str(ipath),
                _parse_cm(term_conf["x"]),
                _parse_cm(term_conf["y"]),
                _parse_cm(term_conf["w"]),
                _parse_cm(term_conf["h"]),
            )

    prs.save(str(img_path))


def _handle_optional_slides(output_path: Path, client_data: dict, brand: str):
    if brand != "vodafone":
        return

    proposal = client_data.get("proposal", client_data)
    include_holidaysim = proposal.get("include_holidaysim", True)
    include_lynks_tic = proposal.get("include_lynks_tic", True)

    if not include_holidaysim:
        try:
            _run_officecli(["remove", str(output_path.resolve()), "/slide[3]"])
        except Exception:
            pass

    if not include_lynks_tic and include_holidaysim:
        try:
            _run_officecli(["remove", str(output_path.resolve()), "/slide[4]"])
        except Exception:
            pass
    elif not include_lynks_tic and not include_holidaysim:
        try:
            _run_officecli(["remove", str(output_path.resolve()), "/slide[3]"])
        except Exception:
            pass

    if proposal.get("include_roaming", False):
        _add_roaming_slides(output_path)


def _add_roaming_slides(output_path: Path):
    roaming_template = BASE_DIR / "templates" / "vodafone_roaming_addon.pptx"
    if not roaming_template.exists():
        return

    roaming_src = str(roaming_template.resolve())
    dst = str(output_path.resolve())

    for slide_num in [3, 4, 5, 6, 7, 8, 9, 10]:
        try:
            _run_officecli(["add", dst, "/", "--type", "slide", "--from", f"{roaming_src}/slide[{slide_num}]"])
        except Exception:
            pass