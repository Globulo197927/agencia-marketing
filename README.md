# Generador de Presentaciones de Precios

Automatiza la generacion de presentaciones comerciales para Vodafone y Lynks-TIC a partir de datos del CRM.

## Estructura

```
Agencia Marketing/
├── config/
│   ├── vodafone.yaml          # Config de marca + mapeo de shapes
│   ├── lynks-tic.yaml         # Config de marca + mapeo de shapes
│   └── settings.yaml          # Config global
├── templates/
│   ├── vodafone_base.pptx      # Template Vodafone (8 slides)
│   ├── vodafone_roaming_addon.pptx  # Slides de roaming (addon)
│   └── lynks-tic_base.pptx    # Template Lynks-TIC (17 slides)
├── input/
│   ├── screenshots/            # Capturas del CRM para OCR
│   └── client_data/            # YAML con datos del cliente
├── output/
│   ├── vodafone/               # Presentaciones generadas
│   └── lynks-tic/
├── overrides/                  # Overrides manuales por cliente
├── src/
│   ├── main.py                 # CLI principal
│   ├── pptx_generator.py       # Generador con officecli
│   ├── ocr_extractor.py        # Extraccion OCR de capturas
│   ├── price_parser.py          # Parseo de precios
│   ├── config_loader.py        # Carga de configs YAML
│   └── models.py               # Dataclasses
└── requirements.txt
```

## Uso

### Generar desde YAML (modo principal)

```bash
# Activar entorno virtual
source .venv/bin/activate

# Vodafone
python src/main.py --brand vodafone --data input/client_data/mi_cliente.yaml

# Lynks-TIC
python src/main.py --brand lynks-tic --data input/client_data/mi_cliente.yaml
```

### Con override manual

```bash
python src/main.py --brand vodafone --data input/client_data/mi_cliente.yaml --override overrides/cliente_vip.yaml
```

### Con OCR (captura del CRM)

```bash
python src/main.py --brand vodafone --screenshot input/screenshots/crm.png --save-yaml input/client_data/extraido.yaml
```

### Incluir slides de roaming (Vodafone)

```bash
python src/main.py --brand vodafone --data input/client_data/mi_cliente.yaml --include-roaming
```

### Solo ver datos (sin generar)

```bash
python src/main.py --brand vodafone --data input/client_data/mi_cliente.yaml --dry-run
```

### Nombre de salida personalizado

```bash
python src/main.py --brand vodafone --data input/client_data/mi_cliente.yaml --output presentacion_cliente_v1
```

## YAML de datos del cliente

### Vodafone (ejemplo)

```yaml
client:
  name_line1: "EMPRESA"
  name_line2: "GLOBAL TECH SL"
  nif: "B87654321"
  lines: 50
  type: "Cliente VIP"

proposal:
  date: "Mayo 2026"
  duration_months: "24 meses"
  network_type: "Red 5G Vodafone"
  include_holidaysim: true
  include_lynks_tic: true
  include_roaming: false

products:
  - name_line1: "TARIFA"
    name_line2: "ILIMITADA 5G"
    price_monthly: "9,50€/mes"
    price_before: "15,00€/mes"
    discount_text: "37% de descuento"
    features:
      - "Datos ilimitados 5G sin limite"
      - "Llamadas ilimitadas en Espana y UE"
      - "Roaming 100 GB incluido"

  - name_line1: "FIBRA"
    name_line2: "600 MBPS"
    price_monthly: "0€/mes"
    price_before: "29,95€/mes"
    discount_text: "100% de descuento"
    features:
      - "Conexion simetrica 600 Mbps"
      - "Incluida sin coste adicional"
      - "Instalacion gratuita"

  - name_line1: "CENTRALITA"
    name_line2: "BASIC"
    price_monthly: "0€/mes"
    price_before: "22,50€/mes"
    discount_text: "100% de descuento"
    features:
      - "Gestion de llamadas basica"
      - "Incluida sin coste adicional"
      - "Extensiones ilimitadas"

economic:
  total_monthly: "475,00 €/mes"
  subsidy_vodafone: "-350,00€"
  subsidy_plan_negocio: "-1.200,00€"
  apoyo_economico: "3.500€"
  summary_line: "50 lineas · Terminales a 0€ · Todos con 37% dto · Compromiso 24 meses"
  subsidy_detail_line: "Sub. Vodafone: -350,00€  ·  Sub. Plan Negocio: -1.200,00€  ·  Apoyo Economico: 3.500€"

terminals:
  - model: "Samsung Galaxy A55 5G"
    price: "0,00 €"
```

### Lynks-TIC (ejemplo)

```yaml
client:
  name_short: "ACME"
  name_suffix: "SERVICIOS PROFESIONALES SL"
  nif: "B12345678"
  sedes: 20
  service_type: "Conectividad FTTO + Voz IP"

proposal:
  date: "Mayo 2026"
  extensiones: 80
  canales: 80
  ddi_unit_price: "1,50€/mes"
  terminal_model: "Yealink W73P"
  terminal_unit_price: "8,50€/ud/mes"

options:
  - name: "OPCION A"
    speed: "1Gbps/1Gbps simetrico"
    speed_download: "1.000 Mbps"
    speed_upload: "1.000 Mbps"
    total_price: "1.890 €/mes"
    ftto_unit_price: "42,00€/mes"
    ftto_total: "840€"
    title: "OPCION A - FTTO 1 GBPS"

  - name: "OPCION B"
    speed: "600Mbps/600Mbps simetrico"
    speed_download: "600 Mbps"
    speed_upload: "600 Mbps"
    total_price: "1.490 €/mes"
    ftto_unit_price: "38,00€/mes"
    ftto_total: "760€"
    title: "OPCION B - FTTO 600 MBPS"
```

## Override manual

Crea un archivo en `overrides/` para sobreescribir campos especificos:

```yaml
# overrides/cliente_vip.yaml
products:
  - price_monthly: "5,50€/mes"    # Precio especial VIP
    discount_text: "50% de descuento"
economic:
  apoyo_economico: "5.000€"      # Apoyo extra
```

## Flujo de trabajo tipico

1. Saca captura del CRM con los precios
2. Opcional: `--screenshot` para extraer via OCR
3. Edita/crea el YAML en `input/client_data/`
4. Ejecuta: `python src/main.py --brand vodafone --data input/client_data/tu_cliente.yaml`
5. La presentacion se genera en `output/vodafone/`
6. Abre el archivo y revisa

## Templates disponibles

| Marca | Template | Slides | Descripcion |
|-------|----------|--------|-------------|
| Vodafone | `vodafone_base.pptx` | 8 | Propuesta completa con productos, pricing, terminales |
| Vodafone | `vodafone_roaming_addon.pptx` | +8 | Slides de roaming por zonas (addon) |
| Lynks-TIC | `lynks-tic_base.pptx` | 17 | Propuesta completa con FTTO, centralita, opciones A/B |

## Dependencias

- Python 3.14+
- officecli (instalado en `~/.local/bin/officecli`)
- PyYAML, Pillow, pytesseract (en `.venv`)
- Tesseract OCR (`brew install tesseract tesseract-lang`)