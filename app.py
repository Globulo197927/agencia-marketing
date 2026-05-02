import streamlit as st
from pathlib import Path
import yaml
import tempfile
import copy

from src.pptx_generator import generate_pptx
from src.ocr_extractor import extract_text_from_image, parse_prices, parse_line_count, parse_client_name
from src.config_loader import BASE_DIR

st.set_page_config(
    page_title="Generador de Propuestas",
    page_icon="📝",
    layout="wide",
)

EXAMPLES = {
    "vodafone": "input/client_data/vodafone_ejemplo.yaml",
    "lynks-tic": "input/client_data/lynks-tic_acme.yaml",
}

def load_example(brand):
    path = BASE_DIR / EXAMPLES[brand]
    return yaml.safe_load(path.read_text())

def data_to_yaml(data):
    return yaml.dump(data, allow_unicode=True, default_flow_style=False, sort_keys=False)

def yaml_to_data(yaml_text):
    if not yaml_text or not yaml_text.strip():
        return None
    try:
        return yaml.safe_load(yaml_text)
    except yaml.YAMLError:
        return None

if "data" not in st.session_state:
    st.session_state["data"] = load_example("vodafone")

brand = st.selectbox(
    "Marca",
    ["vodafone", "lynks-tic"],
    format_func=lambda x: "📱 Vodafone" if x == "vodafone" else "☎️ Lynks-TIC",
    key="brand_select",
)

prev_brand = st.session_state.get("prev_brand", brand)
if prev_brand != brand:
    st.session_state["data"] = load_example(brand)
    st.session_state["prev_brand"] = brand
    st.rerun()
st.session_state["prev_brand"] = brand

st.title("📋 Generador de Propuestas Comerciales")

# ─── STEP 1: UPLOAD CRM SCREENSHOT ───
with st.expander("📸 1. Subir captura del CRM (opcional)", expanded=False):
    st.caption("Sube una captura del CRM para extraer datos automáticamente. Después revisa y edita abajo.")
    uploaded = st.file_uploader("Arrastra la captura aquí", type=["png", "jpg", "jpeg"])

    if uploaded:
        with st.spinner("Extrayendo datos..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                img_path = Path(tmpdir) / uploaded.name
                with open(img_path, "wb") as f:
                    f.write(uploaded.getbuffer())
                raw_text = extract_text_from_image(str(img_path))
                prices = parse_prices(raw_text)
                line_info = parse_line_count(raw_text)
                client_info = parse_client_name(raw_text)

                col1, col2 = st.columns(2)
                with col1:
                    with st.expander("Texto OCR", expanded=False):
                        st.text(raw_text[:3000])
                with col2:
                    st.json({"cliente": client_info, "líneas": line_info, "precios": len(prices)})

                data = st.session_state["data"]
                if brand == "vodafone":
                    if client_info.get("company_name"):
                        parts = client_info["company_name"].strip().rsplit(" ", 1)
                        data.setdefault("client", {})
                        data["client"]["name_line1"] = parts[0] if len(parts) > 1 else client_info["company_name"]
                        data["client"]["name_line2"] = parts[1] if len(parts) > 1 else ""
                    if client_info.get("nif"):
                        data.setdefault("client", {})
                        data["client"]["nif"] = client_info["nif"]
                    if line_info.get("lines"):
                        data.setdefault("client", {})
                        data["client"]["lines"] = line_info["lines"]
                elif brand == "lynks-tic":
                    if client_info.get("company_name"):
                        parts = client_info["company_name"].strip().rsplit(" ", 1)
                        data.setdefault("client", {})
                        data["client"]["name_short"] = parts[0] if len(parts) > 1 else client_info["company_name"]
                        data["client"]["name_suffix"] = parts[1] if len(parts) > 1 else ""
                    if client_info.get("nif"):
                        data.setdefault("client", {})
                        data["client"]["nif"] = client_info["nif"]
                    if line_info.get("lines"):
                        data.setdefault("proposal", {})
                        data["proposal"]["extensiones"] = line_info["lines"]

                st.session_state["data"] = data
                st.success("Datos extraídos y aplic al formulario de abajo.")
                st.rerun()

st.divider()

# ─── STEP 2: EDIT FIELDS ───
data = copy.deepcopy(st.session_state["data"])

if brand == "vodafone":
    st.subheader("2. Datos del cliente y propuesta (Vodafone)")

    tab_client, tab_proposal, tab_products, tab_economic, tab_terminals, tab_yaml = st.tabs(
        ["👤 Cliente", "📋 Propuesta", "📦 Productos", "💰 Económicos", "📱 Terminales", " YAML"]
    )

    with tab_client:
        c = data.setdefault("client", {})
        c1, c2 = st.columns(2)
        with c1:
            c["name_line1"] = st.text_input("Nombre línea 1", value=c.get("name_line1", ""), key="v_nl1")
            c["nif"] = st.text_input("NIF", value=c.get("nif", ""), key="v_nif")
            c["lines"] = st.number_input("Nº de líneas", min_value=1, value=int(c.get("lines", 25)), key="v_lines")
        with c2:
            c["name_line2"] = st.text_input("Nombre línea 2", value=c.get("name_line2", ""), key="v_nl2")
            c["type"] = st.selectbox("Tipo cliente", ["Cliente Nuevo", "Cliente Existente", "Portabilidad"],
                                      index=["Cliente Nuevo", "Cliente Existente", "Portabilidad"].index(c.get("type", "Cliente Nuevo")),
                                      key="v_ctype")

    with tab_proposal:
        p = data.setdefault("proposal", {})
        c1, c2 = st.columns(2)
        with c1:
            p["date"] = st.text_input("Fecha", value=p.get("date", "Mayo 2026"), key="v_date")
            p["duration_months"] = st.text_input("Duración", value=str(p.get("duration_months", "36 meses")), key="v_dur")
        with c2:
            p["network_type"] = st.text_input("Tipo de red", value=p.get("network_type", "Red 5G Vodafone"), key="v_net")
        c1, c2, c3 = st.columns(3)
        with c1:
            p["include_holidaysim"] = st.checkbox("HolidaySim", value=p.get("include_holidaysim", True), key="v_hol")
        with c2:
            p["include_lynks_tic"] = st.checkbox("Lynks-TIC", value=p.get("include_lynks_tic", True), key="v_lynks")
        with c3:
            p["include_roaming"] = st.checkbox("Roaming", value=p.get("include_roaming", False), key="v_roam")

    with tab_products:
        products = data.setdefault("products", [])
        defaults = [
            {"name_line1": "BUNDLE RED", "name_line2": "INFINITY PRO", "price_monthly": "73,44€/mes", "price_before": "112,98€/mes", "discount_text": "35% de descuento",
             "features": ["Datos ilimitados Full Speed", "5G en toda la red Vodafone", "Roaming incluido en UE"]},
            {"name_line1": "FIBRA", "name_line2": "1 GBPS", "price_monthly": "0€/mes", "price_before": "34,85€/mes", "discount_text": "100% de descuento",
             "features": ["Conexión simétrica 1 Gbps", "Incluida sin coste adicional", "Instalación y mantenimiento"]},
            {"name_line1": "CENTRALITA", "name_line2": "PLUS", "price_monthly": "0€/mes", "price_before": "30,75€/mes", "discount_text": "100% de descuento",
             "features": ["Gestión avanzada de llamadas", "Incluida sin coste adicional", "Multi-sede y extensiones"]},
        ]
        while len(products) < 3:
            products.append(defaults[len(products)] if len(products) < len(defaults) else {})

        num_prods = st.number_input("Nº de productos", min_value=1, max_value=3, value=len(products), key="v_nprods")
        products_trimmed = products[:num_prods]
        for i, prod in enumerate(products_trimmed):
            def_prod = defaults[i] if i < len(defaults) else {}
            with st.expander(f"Producto {i+1}: {(prod.get('name_line1','') + ' ' + prod.get('name_line2','')).strip() or f'Producto {i+1}'}", expanded=(i==0)):
                pc1, pc2 = st.columns(2)
                with pc1:
                    prod["name_line1"] = st.text_input("Nombre línea 1", value=prod.get("name_line1", def_prod.get("name_line1", "")), key=f"v_p{i}_n1")
                    prod["price_monthly"] = st.text_input("Precio/mes", value=prod.get("price_monthly", def_prod.get("price_monthly", "")), key=f"v_p{i}_pm")
                    prod["discount_text"] = st.text_input("Descuento", value=prod.get("discount_text", def_prod.get("discount_text", "")), key=f"v_p{i}_dt")
                with pc2:
                    prod["name_line2"] = st.text_input("Nombre línea 2", value=prod.get("name_line2", def_prod.get("name_line2", "")), key=f"v_p{i}_n2")
                    prod["price_before"] = st.text_input("Precio antes", value=prod.get("price_before", def_prod.get("price_before", "")), key=f"v_p{i}_pb")
                feats = prod.get("features", def_prod.get("features", []))
                feats_text = st.text_area("Características (una por línea)", value="\n".join(feats), key=f"v_p{i}_feat")
                prod["features"] = [f.strip() for f in feats_text.split("\n") if f.strip()]
        data["products"] = products_trimmed

    with tab_economic:
        eco = data.setdefault("economic", {})
        c1, c2 = st.columns(2)
        with c1:
            eco["total_monthly"] = st.text_input("Total mensual", value=eco.get("total_monthly", ""), key="v_tm")
            eco["apoyo_economico"] = st.text_input("Apoyo económico", value=eco.get("apoyo_economico", ""), key="v_ae")
        with c2:
            eco["summary_line"] = st.text_input("Resumen", value=eco.get("summary_line", ""), key="v_sl")
            eco["subsidy_detail_line"] = st.text_input("Detalle subsidios", value=eco.get("subsidy_detail_line", ""), key="v_sd")

        dist = data.setdefault("distribution_table", [])
        st.markdown("**Tabla de distribución**")
        st.caption("Deja vacío para usar valores por defecto de la plantilla.")
        default_dist = [
            {"terminal": "Adoc NEO 8000 4G Deals", "plan": "RED Infinity PRO Ilim. Bundle", "precio_cesion": "187,00€", "sin_dto": "0,00€", "dto": "35%", "con_dto": "0,00€", "uds": "1", "total": "0,00€"},
            {"terminal": "Adoc NEO 3850W 4G Deals", "plan": "RED Infinity PRO Ilim. Full Speed", "precio_cesion": "365,00€", "sin_dto": "17,95€", "dto": "35%", "con_dto": "11,67€", "uds": "5", "total": "58,34€"},
            {"terminal": "Adoc NEO 3850W 4G Deals", "plan": "RED Infinity PRO Ilim. Full Speed", "precio_cesion": "365,00€", "sin_dto": "17,95€", "dto": "35%", "con_dto": "11,67€", "uds": "5", "total": "58,34€"},
            {"terminal": "—", "plan": "RED Infinity PRO 5Gb", "precio_cesion": "0,00€", "sin_dto": "13,73€", "dto": "35%", "con_dto": "8,92€", "uds": "10", "total": "89,24€"},
            {"terminal": "—", "plan": "RED Infinity PRO 5Gb", "precio_cesion": "0,00€", "sin_dto": "13,73€", "dto": "35%", "con_dto": "8,92€", "uds": "1", "total": "8,92€"},
            {"terminal": "Cocomm DT250 4G Solo Voz", "plan": "RED Infinity PRO Solo Voz", "precio_cesion": "198,00€", "sin_dto": "8,98€", "dto": "35%", "con_dto": "5,84€", "uds": "3", "total": "17,51€"},
        ]
        while len(dist) < 6:
            dist.append(default_dist[len(dist)])
        for i, row in enumerate(dist[:6]):
            with st.expander(f"Fila {i+1}: {row.get('terminal', '—')}", expanded=False):
                drow = default_dist[i] if i < len(default_dist) else {}
                rc1, rc2 = st.columns(2)
                with rc1:
                    row["terminal"] = st.text_input("Terminal", value=row.get("terminal", drow.get("terminal", "")), key=f"v_d{i}_t")
                    row["precio_cesion"] = st.text_input("Precio cesión", value=row.get("precio_cesion", drow.get("precio_cesion", "")), key=f"v_d{i}_pc")
                    row["dto"] = st.text_input("Descuento", value=row.get("dto", drow.get("dto", "")), key=f"v_d{i}_dt")
                    row["total"] = st.text_input("Total", value=row.get("total", drow.get("total", "")), key=f"v_d{i}_tot")
                with rc2:
                    row["plan"] = st.text_input("Plan", value=row.get("plan", drow.get("plan", "")), key=f"v_d{i}_p")
                    row["sin_dto"] = st.text_input("Sin dto", value=row.get("sin_dto", drow.get("sin_dto", "")), key=f"v_d{i}_sd")
                    row["con_dto"] = st.text_input("Con dto", value=row.get("con_dto", drow.get("con_dto", "")), key=f"v_d{i}_cd")
                    row["uds"] = st.text_input("Uds", value=row.get("uds", drow.get("uds", "")), key=f"v_d{i}_u")
        data["distribution_table"] = dist[:6]

    with tab_terminals:
        terms = data.setdefault("terminals", [])
        default_terms = [
            {"model": "Adoc NEO 8000 4G", "price": "0,00 €"},
            {"model": "Adoc NEO 3850W 4G", "price": "0,00 €"},
            {"model": "Cocomm DT250 4G Solo Voz", "price": "0,00 €"},
        ]
        while len(terms) < 3:
            terms.append(default_terms[len(terms)])
        num_terms = st.number_input("Nº de terminales", min_value=0, max_value=3, value=len(terms), key="v_nterm")
        for i, t in enumerate(terms[:num_terms]):
            dt = default_terms[i] if i < len(default_terms) else {}
            with st.expander(f"Terminal {i+1}", expanded=(i==0)):
                t["model"] = st.text_input("Modelo", value=t.get("model", dt.get("model", "")), key=f"v_t{i}_m")
                t["price"] = st.text_input("Precio cesión", value=t.get("price", dt.get("price", "0,00 €")), key=f"v_t{i}_p")
        data["terminals"] = terms[:num_terms]

    with tab_yaml:
        yaml_text = st.text_area("YAML (editable)", value=data_to_yaml(data), height=500, key="v_yaml")
        yaml_data = yaml_to_data(yaml_text)
        if yaml_data:
            data = yaml_data
            st.session_state["data"] = data

# ─── LYNKS-TIC ───
else:
    st.subheader("2. Datos del cliente y propuesta (Lynks-TIC)")

    tab_client, tab_proposal, tab_options, tab_yaml = st.tabs(
        ["👤 Cliente", "📋 Propuesta", "📦 Opciones", " YAML"]
    )

    data = copy.deepcopy(st.session_state["data"])

    with tab_client:
        c = data.setdefault("client", {})
        c1, c2 = st.columns(2)
        with c1:
            c["name_short"] = st.text_input("Nombre corto", value=c.get("name_short", ""), key="l_ns")
            c["nif"] = st.text_input("NIF", value=c.get("nif", ""), key="l_nif")
            c["sedes"] = st.number_input("Nº sedes", min_value=1, value=int(c.get("sedes", 1)), key="l_sedes")
        with c2:
            c["name_suffix"] = st.text_input("Sufijo nombre", value=c.get("name_suffix", ""), key="l_nsu")
            c["service_type"] = st.text_input("Tipo servicio", value=c.get("service_type", "Conectividad FTTO + Voz IP"), key="l_svc")

    with tab_proposal:
        p = data.setdefault("proposal", {})
        c1, c2 = st.columns(2)
        with c1:
            p["date"] = st.text_input("Fecha", value=p.get("date", "Mayo 2026"), key="l_date")
            p["extensiones"] = st.number_input("Extensiones", min_value=1, value=int(p.get("extensiones", 1)), key="l_ext")
        with c2:
            p["canales"] = st.number_input("Canales", min_value=1, value=int(p.get("canales", 1)), key="l_can")
            p["terminal_model"] = st.text_input("Modelo terminal", value=p.get("terminal_model", "Yealink W73P"), key="l_mt")
        p["terminal_unit_price"] = st.text_input("Precio terminal/ud/mes", value=p.get("terminal_unit_price", "8,50€/ud/mes"), key="l_tp")

    with tab_options:
        opts = data.setdefault("options", [])
        defaults_opts = [
            {"title": "OPCIÓN A — FTTO 1 GBPS · LA MÁS POTENTE", "speed": "1Gbps/1Gbps simétrico", "total_price": "1.890 €/mes"},
            {"title": "OPCIÓN B — FTTO 600 MBPS · LA EQUILIBRADA", "speed": "600Mbps/600Mbps simétrico", "total_price": "1.490 €/mes"},
        ]
        while len(opts) < 2:
            opts.append(defaults_opts[len(opts)])
        for i, opt in enumerate(opts[:2]):
            dopt = defaults_opts[i]
            with st.expander(f"Opción {chr(65+i)}", expanded=True):
                opt["title"] = st.text_input("Título", value=opt.get("title", dopt["title"]), key=f"l_o{i}_ti")
                opt["total_price"] = st.text_input("Precio total", value=opt.get("total_price", dopt["total_price"]), key=f"l_o{i}_tp")
                opt["speed"] = st.text_input("Velocidad", value=opt.get("speed", dopt["speed"]), key=f"l_o{i}_sp")
        data["options"] = opts[:2]
        data["savings_line"] = st.text_input("Línea ahorro", value=data.get("savings_line", ""), key="l_sv")

    with tab_yaml:
        yaml_text = st.text_area("YAML (editable)", value=data_to_yaml(data), height=500, key="l_yaml")
        yaml_data = yaml_to_data(yaml_text)
        if yaml_data:
            data = yaml_data
            st.session_state["data"] = data

st.session_state["data"] = data

st.divider()

# ─── STEP 3: GENERATE ───
st.subheader("3. Generar propuesta")

col_gen, col_example = st.columns([3, 1])
with col_example:
    if st.button("📄 Cargar ejemplo"):
        st.session_state["data"] = load_example(brand)
        st.rerun()

with col_gen:
    if st.button("🚀 Generar PPTX", type="primary", use_container_width=True):
        try:
            output = generate_pptx(brand, data)
            st.success("✅ Propuesta generada correctamente")

            with open(output, "rb") as f:
                st.download_button(
                    "📥 Descargar PPTX",
                    data=f.read(),
                    file_name=output.name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
        except Exception as e:
            st.error(f"Error al generar: {e}")