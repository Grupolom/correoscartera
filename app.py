import os
import json
import resend
import webbrowser
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, render_template, request, jsonify
from flask_cors import CORS
from threading import Timer
from dotenv import load_dotenv
from datetime import datetime, date, timedelta
import pandas as pd
from io import BytesIO


# Cargar variables de entorno desde .env
load_dotenv()


app = Flask(__name__)
CORS(app)


# ==========================================
# CONFIGURACIÓN DE CORREO ELECTRÓNICO
# ==========================================
RESEND_API_KEY = os.getenv("RESEND_API_KEY", "")
EMAIL_FROM_NAME = os.getenv("EMAIL_FROM_NAME", "Cartera Lomarosa")
EMAIL_FROM_ADDRESS = os.getenv("EMAIL_FROM_ADDRESS", "cartera@grupolom.com")

resend.api_key = RESEND_API_KEY


MAX_WORKERS = int(os.getenv("MAX_WORKERS", "3"))


# ==========================================
# CORREOS FIJOS DE CC
# ==========================================

CC_CARTERA = "cartera@grupolom.com"

CORREOS_VENDEDORES = {
    "padilla manga william dario":          "w.padilla@grupolom.com",
    "pardo prada jorge alberto":            "j.pardo@grupolom.com",
    "silvestre acosta angela lucia":        "a.silvestre@grupolom.com",
    "gonzalez triana jennifer":             "comercial.lomarosa@gmail.com",
    "de jesus oropeza yohn robinson":       "supervisortat@grupolom.com",
    "rodriguez rodriguez wilson javier":    "j.rodriguez@grupolom.com",
    "farfan ospina juan carlos":            "asesorcomercial@grupolom.com",
    "pardo prada alfredo":                  "j.pardo@grupolom.com",
}


# ==========================================
# FUNCIONES DE NORMALIZACIÓN
# ==========================================


def normalizar_nombre(nombre):
    if not nombre:
        return ""
    return str(nombre).strip().lower()


def normalizar_columna(col):
    return str(col).strip().lower().replace('  ', ' ')


def normalizar_nit(nit_str):
    """Extrae solo dígitos del NIT para comparación robusta."""
    if not nit_str:
        return None
    digits = ''.join(filter(str.isdigit, str(nit_str)))
    return digits if digits else None


# ==========================================
# FUNCIONES DE FECHA CARTERA Y CIERRE TRIMESTRAL
# ==========================================


def extraer_fecha_cartera(archivo_bytes):
    """Extrae la fecha de corte del Excel de cartera. Soporta formato nuevo y antiguo."""
    try:
        xl = pd.ExcelFile(BytesIO(archivo_bytes))

        # Formato nuevo: "Informe de Vencimientos" — busca "Fecha de Corte:" en primeras filas
        if "Informe de Vencimientos" in xl.sheet_names:
            df_raw = pd.read_excel(BytesIO(archivo_bytes), sheet_name="Informe de Vencimientos", header=None, nrows=12)
            for idx in range(len(df_raw)):
                for col_idx in range(len(df_raw.columns)):
                    val = df_raw.iloc[idx, col_idx]
                    if pd.notna(val) and "fecha de corte" in str(val).lower():
                        val_str = str(val).strip()
                        if ":" in val_str:
                            date_part = val_str.split(":", 1)[1].strip()
                            try:
                                fecha = pd.to_datetime(date_part).date()
                                print(f"[INFO] Fecha de Corte encontrada: {fecha.strftime('%d/%m/%Y')}")
                                return fecha
                            except:
                                pass
                        # Intentar columna adyacente
                        if col_idx + 1 < len(df_raw.columns):
                            next_val = df_raw.iloc[idx, col_idx + 1]
                            if pd.notna(next_val):
                                try:
                                    fecha = pd.to_datetime(next_val).date()
                                    print(f"[INFO] Fecha de Corte encontrada: {fecha.strftime('%d/%m/%Y')}")
                                    return fecha
                                except:
                                    pass
            print("[WARNING] No se encontró 'Fecha de Corte' en 'Informe de Vencimientos'")
            return date.today()

        # Formato antiguo: "Cartera por edades" — busca "Fecha Cartera" en columna E
        if "Cartera por edades" in xl.sheet_names:
            df_raw = pd.read_excel(BytesIO(archivo_bytes), sheet_name="Cartera por edades", header=None)
            for idx, valor in enumerate(df_raw.iloc[:, 4]):
                if pd.notna(valor) and "fecha cartera" in str(valor).lower():
                    fecha_raw = df_raw.iloc[idx, 5]
                    if pd.notna(fecha_raw):
                        try:
                            fecha = pd.to_datetime(fecha_raw).date()
                            print(f"[INFO] Fecha Cartera encontrada: {fecha.strftime('%d/%m/%Y')}")
                            return fecha
                        except:
                            pass
                    break

        print("[WARNING] No se detectó fecha de cartera. Usando fecha de hoy.")
        return date.today()

    except Exception as e:
        print(f"[ERROR] Error al extraer fecha cartera: {e}")
        return date.today()


def detectar_cierre_trimestral(fecha_cartera):
    """
    Detecta si la fecha corresponde a un cierre trimestral.
    Retorna dict con información del cierre o None si no aplica.

    Cierres: días 29, 30, 31 de Marzo, Junio, Septiembre, Diciembre
    """
    if not fecha_cartera:
        return None

    mes = fecha_cartera.month
    dia = fecha_cartera.day

    # Meses de cierre trimestral
    meses_cierre = {
        3: "trimestral",   # Marzo
        6: "trimestral",   # Junio
        9: "trimestral",   # Septiembre
        12: "anual"        # Diciembre
    }

    # Días válidos para cierre
    dias_cierre = [29, 30, 31]

    if mes in meses_cierre and dia in dias_cierre:
        tipo_cierre = meses_cierre[mes]
        return {
            "es_cierre": True,
            "tipo": tipo_cierre,  # "trimestral" o "anual"
            "fecha": fecha_cartera,
            "fecha_formateada": fecha_cartera.strftime("%d/%m/%Y"),
            "mensaje_tipo": "Cierre anual" if tipo_cierre == "anual" else "Cierre trimestral"
        }

    return None


# ==========================================
# FUNCIONES DE AGRUPACIÓN
# ==========================================


def agrupar_recordatorios_por_cliente(recordatorios):
    """
    Agrupa recordatorios por cliente+email (sin separar por estado).

    Retorna una estructura unificada con:
    - facturas_vencidas[]
    - facturas_proximas[]
    - facturas_no_vencidas[]
    - métricas agregadas
    """
    agrupados = {}

    for recordatorio in recordatorios:
        cliente_nombre = recordatorio.get("cliente")
        cliente_email = recordatorio.get("correo_cliente")
        estado = recordatorio.get("estado")

        # Key único: cliente + email (UN SOLO correo por cliente)
        key = f"{cliente_nombre}|{cliente_email}"

        if key not in agrupados:
            agrupados[key] = {
                "cliente": cliente_nombre,
                "correo_cliente": cliente_email,
                "vendedor": recordatorio.get("vendedor"),
                "correo_vendedor": recordatorio.get("correo_vendedor"),
                "local": recordatorio.get("local"),
                "facturas_vencidas": [],
                "facturas_proximas": [],
                "facturas_no_vencidas": [],
                "total_facturas": 0,
                "total_vencidas": 0,
                "total_proximas": 0,
                "total_no_vencidas": 0,
                "total_saldo": 0,
                "cupo": recordatorio.get("cupo", 0),
                "cupo_disponible": 0  # Se calcula al final
            }

        # Construir objeto de factura
        factura_obj = {
            "numero_factura": recordatorio.get("numero_factura"),
            "fecha_emision": recordatorio.get("fecha_emision"),
            "fecha_vencimiento": recordatorio.get("fecha_vencimiento"),
            "dias": recordatorio.get("dias"),
            "saldo": recordatorio.get("saldo"),
            "saldo_numerico": recordatorio.get("saldo_numerico"),
            "estado": estado
        }

        # Clasificar en array correspondiente
        if estado == "vencido":
            agrupados[key]["facturas_vencidas"].append(factura_obj)
            agrupados[key]["total_vencidas"] += 1
        elif estado == "proximo":
            agrupados[key]["facturas_proximas"].append(factura_obj)
            agrupados[key]["total_proximas"] += 1
        elif estado == "no_vencido":
            agrupados[key]["facturas_no_vencidas"].append(factura_obj)
            agrupados[key]["total_no_vencidas"] += 1

        # Actualizar métricas generales
        agrupados[key]["total_facturas"] += 1
        agrupados[key]["total_saldo"] += recordatorio.get("saldo_numerico", 0)

    # Calcular cupo_disponible para cada cliente
    for cliente in agrupados.values():
        cliente["cupo_disponible"] = cliente["cupo"] - cliente["total_saldo"]

    resultado = list(agrupados.values())

    print(f"\n[INFO] Agrupación unificada por cliente + email:")
    print(f"  - Recordatorios individuales (facturas): {len(recordatorios)}")
    print(f"  - Clientes únicos a notificar: {len(resultado)}")

    total_vencidas = sum(c["total_vencidas"] for c in resultado)
    total_proximas = sum(c["total_proximas"] for c in resultado)
    total_no_vencidas = sum(c["total_no_vencidas"] for c in resultado)

    print(f"    • Total facturas vencidas: {total_vencidas}")
    print(f"    • Total facturas próximas: {total_proximas}")
    print(f"    • Total facturas no vencidas: {total_no_vencidas}")
    print(f"  - Nota: Cada cliente recibirá UN SOLO correo con todas sus facturas")

    return resultado




def dividir_en_lotes(recordatorios, limite=450):
    """Divide los recordatorios en lotes de máximo 'limite' correos."""
    lote1 = recordatorios[:limite]
    lote2 = recordatorios[limite:]
    return lote1, lote2


# ==========================================
# FUNCIONES DE LECTURA DE EXCEL
# ==========================================


def _detectar_sheet_cartera(archivo_bytes):
    """Retorna el nombre de la hoja de cartera del archivo, o None si no es un archivo de cartera."""
    try:
        xl = pd.ExcelFile(BytesIO(archivo_bytes))
        if "Informe de Vencimientos" in xl.sheet_names:
            return "Informe de Vencimientos"
        if "Cartera por edades" in xl.sheet_names:
            return "Cartera por edades"
    except:
        pass
    return None


def detectar_fila_header_cartera(archivo_bytes):
    """Detecta dinámicamente la fila del header en el archivo de cartera."""
    try:
        xl = pd.ExcelFile(BytesIO(archivo_bytes))

        # Formato nuevo
        if "Informe de Vencimientos" in xl.sheet_names:
            df_raw = pd.read_excel(BytesIO(archivo_bytes), sheet_name="Informe de Vencimientos", header=None, nrows=15)
            for idx, valor in enumerate(df_raw.iloc[:, 0]):
                if pd.notna(valor) and str(valor).strip().lower() == "cuenta":
                    print(f"[DEBUG] Header cartera (nuevo formato) en fila {idx}")
                    return idx
            return 9

        # Formato antiguo
        if "Cartera por edades" in xl.sheet_names:
            df_raw = pd.read_excel(BytesIO(archivo_bytes), sheet_name="Cartera por edades", header=None, nrows=20)
            for idx, valor in enumerate(df_raw.iloc[:, 0]):
                if pd.notna(valor) and "nombre tercero" in str(valor).lower():
                    print(f"[DEBUG] Header cartera (formato antiguo) en fila {idx}")
                    return idx
            return 11

    except Exception as e:
        print(f"[DEBUG] Error buscando header: {e}")
    return 9


def detectar_tipo_excel(df, debug_info=""):
    """Detecta si el DataFrame es de Clientes o de Cartera. Soporta formato nuevo y antiguo."""
    columnas_lower = [normalizar_columna(col) for col in df.columns]
    columnas_str = " ".join(columnas_lower)

    print(f"[DEBUG] Detectando tipo {debug_info} | columnas: {columnas_lower[:10]}")

    # --- CLIENTES ---
    tiene_nit = "nit" in columnas_str
    tiene_tercero = "tercero" in columnas_str
    tiene_cliente = "cliente" in columnas_str
    tiene_email = "email" in columnas_str
    tiene_correo_cliente = "correo cliente" in columnas_str or "correocliente" in columnas_str.replace(' ', '')

    # --- CARTERA formato nuevo: Cuenta / Documento / Fecha Vence / Días Vence / Saldo ---
    tiene_cuenta = "cuenta" in columnas_str
    tiene_documento = "documento" in columnas_str
    tiene_fecha_vence = "fecha vence" in columnas_str or "fechavence" in columnas_str.replace(' ', '')
    tiene_dias_vence = "días vence" in columnas_str or "dias vence" in columnas_str

    # --- CARTERA formato antiguo: Nombre tercero / Numero FAC / Vencimiento / Dias / Saldo ---
    tiene_nombre_tercero = "nombre tercero" in columnas_str or "nombretercero" in columnas_str.replace(' ', '')
    tiene_numero_fac = "numero fac" in columnas_str or "numerofac" in columnas_str.replace(' ', '') or " fac " in columnas_str
    tiene_vencimiento = "vencimiento" in columnas_str
    tiene_dias = "dias" in columnas_str or "días" in columnas_str
    tiene_saldo = "saldo" in columnas_str

    if tiene_nit and (tiene_tercero or tiene_cliente) and (tiene_email or tiene_correo_cliente):
        print(f"[DEBUG] [OK] CLIENTES {debug_info}")
        return "clientes"
    elif tiene_cuenta and tiene_documento and tiene_fecha_vence and tiene_dias_vence and tiene_saldo:
        print(f"[DEBUG] [OK] CARTERA nuevo formato {debug_info}")
        return "cartera"
    elif tiene_nombre_tercero and tiene_numero_fac and tiene_vencimiento and tiene_dias and tiene_saldo:
        print(f"[DEBUG] [OK] CARTERA formato antiguo {debug_info}")
        return "cartera"
    else:
        print(f"[DEBUG] [NO DETECTADO] {debug_info}")
        return None


def buscar_columna_exacta(df, nombres_esperados):
    """Busca una columna en el DataFrame con nombres esperados (flexible con espacios)."""
    columnas_map = {normalizar_columna(col): col for col in df.columns}
    
    for nombre_esperado in nombres_esperados:
        nombre_norm = normalizar_columna(nombre_esperado)
        
        if nombre_norm in columnas_map:
            return columnas_map[nombre_norm]
        
        nombre_sin_espacios = nombre_norm.replace(' ', '')
        for col_norm, col_original in columnas_map.items():
            if nombre_sin_espacios == col_norm.replace(' ', ''):
                return col_original
        
        for col_norm, col_original in columnas_map.items():
            if nombre_norm in col_norm or nombre_sin_espacios in col_norm.replace(' ', ''):
                return col_original
    
    return None


def leer_excel_clientes(archivo_bytes):
    """Lee Excel de Clientes. Soporta formato nuevo (Tercero/Email, header fila 1) y antiguo."""

    # Detectar si tiene fila decorativa en fila 0 → header real en fila 1
    df_peek = pd.read_excel(BytesIO(archivo_bytes), header=0, nrows=1)
    cols_peek = [str(c).lower() for c in df_peek.columns]
    if any("exported" in c or (c.startswith("unnamed") and i == 0) for i, c in enumerate(cols_peek)):
        df = pd.read_excel(BytesIO(archivo_bytes), header=1)
        print("[INFO] Excel Clientes: header en fila 1 (formato nuevo)")
    else:
        df = pd.read_excel(BytesIO(archivo_bytes))
        print("[INFO] Excel Clientes: header en fila 0 (formato antiguo)")

    print(f"[DEBUG] Columnas detectadas: {list(df.columns)}")

    col_nit             = buscar_columna_exacta(df, ["Nit", "NIT"])
    col_tercero         = buscar_columna_exacta(df, ["Tercero", "Cliente", "tercero", "cliente"])
    col_email           = buscar_columna_exacta(df, ["Email", "Correo cliente", "Correo", "email"])
    col_vendedor        = buscar_columna_exacta(df, ["Vendedor", "vendedor"])
    col_correo_vendedor = buscar_columna_exacta(df, ["Correo vendedor", "Correovendedor", "Email vendedor"])
    col_cupo            = buscar_columna_exacta(df, ["Cupo Asignado", "Cupo", "Cupo de crédito", "Cupo de credito", "Cupo credito"])

    if not col_tercero:
        raise ValueError(f"No se encontró columna de cliente/tercero. Columnas: {list(df.columns)}")
    if not col_email:
        raise ValueError(f"No se encontró columna de email. Columnas: {list(df.columns)}")

    print(f"[INFO] Mapeo de columnas:")
    print(f"  - Tercero/Cliente : {col_tercero}")
    print(f"  - Email           : {col_email}")
    print(f"  - Vendedor        : {col_vendedor}")
    print(f"  - Correo vendedor : {col_correo_vendedor if col_correo_vendedor else '[No disponible - sin CC]'}")
    print(f"  - Cupo            : {col_cupo if col_cupo else '[No encontrado - $0 por defecto]'}")

    dict_clientes = {}
    dict_vendedores = {}

    for _, row in df.iterrows():
        tercero = row[col_tercero] if pd.notna(row[col_tercero]) else None
        email   = row[col_email]   if pd.notna(row[col_email])   else None

        if not tercero or not email:
            continue

        tercero_norm = normalizar_nombre(tercero)
        if not tercero_norm:
            continue

        # Normalizar NIT (solo dígitos) para indexar por NIT
        nit_clean = None
        if col_nit and pd.notna(row[col_nit]):
            try:
                nit_clean = normalizar_nit(str(int(float(row[col_nit]))))
            except:
                nit_clean = normalizar_nit(str(row[col_nit]))

        # Solo primera ocurrencia por nombre (primera sucursal en formato nuevo)
        if tercero_norm not in dict_clientes:
            cupo_valor = 0
            if col_cupo and pd.notna(row[col_cupo]):
                try:
                    cupo_valor = float(row[col_cupo])
                except:
                    cupo_valor = 0

            vendedor_nombre = str(row[col_vendedor]).strip() if col_vendedor and pd.notna(row[col_vendedor]) else "N/A"
            correo_vendedor = CORREOS_VENDEDORES.get(normalizar_nombre(vendedor_nombre), "N/A")

            entrada = {
                "nit": nit_clean or "N/A",
                "cliente": str(tercero).strip(),
                "correo_cliente": str(email).strip(),
                "vendedor": vendedor_nombre,
                "correo_vendedor": correo_vendedor,
                "cupo": cupo_valor
            }
            dict_clientes[tercero_norm] = entrada

            # Índice secundario por NIT para matching desde cartera
            if nit_clean:
                dict_clientes[f"__nit__{nit_clean}"] = entrada

        # Vendedores con correo (formato antiguo)
        if col_correo_vendedor:
            vendedor = row[col_vendedor] if col_vendedor and pd.notna(row[col_vendedor]) else None
            correo_v = row[col_correo_vendedor] if pd.notna(row[col_correo_vendedor]) else None
            if vendedor and correo_v:
                vend_norm = normalizar_nombre(vendedor)
                if vend_norm:
                    dict_vendedores[vend_norm] = str(correo_v).strip()
                    if tercero_norm in dict_clientes:
                        dict_clientes[tercero_norm]["correo_vendedor"] = str(correo_v).strip()

    clientes_reales = [k for k in dict_clientes if not k.startswith("__nit__")]
    print(f"[INFO] Excel Clientes procesado: {len(clientes_reales)} clientes, {len(dict_vendedores)} vendedores con correo")

    return dict_clientes, dict_vendedores


def leer_excel_cartera(archivo_bytes, dict_clientes, dict_vendedores):
    """Lee Excel de Cartera. Detecta automáticamente el formato (nuevo o antiguo)."""
    sheet = _detectar_sheet_cartera(archivo_bytes)
    if sheet == "Informe de Vencimientos":
        return _leer_cartera_nuevo_formato(archivo_bytes, dict_clientes, dict_vendedores)
    elif sheet == "Cartera por edades":
        return _leer_cartera_formato_antiguo(archivo_bytes, dict_clientes, dict_vendedores)
    else:
        raise ValueError(f"No se reconoce ninguna hoja de cartera en el archivo.")


def _leer_cartera_nuevo_formato(archivo_bytes, dict_clientes, dict_vendedores):
    """Parsea el nuevo formato de cartera: hoja 'Informe de Vencimientos'."""
    header_row = detectar_fila_header_cartera(archivo_bytes)
    print(f"[INFO] Leyendo 'Informe de Vencimientos' con header en fila {header_row}")
    df = pd.read_excel(BytesIO(archivo_bytes), sheet_name="Informe de Vencimientos", header=header_row)

    col_cuenta     = buscar_columna_exacta(df, ["Cuenta", "cuenta"])
    col_documento  = buscar_columna_exacta(df, ["Documento", "documento"])
    col_fecha      = buscar_columna_exacta(df, ["Fecha", "Emision", "Emisión"])
    col_fecha_vence= buscar_columna_exacta(df, ["Fecha Vence", "FechaVence", "Vencimiento", "Fecha Vencimiento"])
    col_saldo      = buscar_columna_exacta(df, ["Saldo", "saldo"])

    if not col_cuenta or not col_documento or not col_fecha_vence or not col_saldo:
        raise ValueError(f"Columnas requeridas no encontradas. Columnas disponibles: {list(df.columns)}")

    print(f"[INFO] Columnas cartera nuevo formato:")
    print(f"  - Cuenta     : {col_cuenta}")
    print(f"  - Documento  : {col_documento}")
    print(f"  - Fecha      : {col_fecha}")
    print(f"  - Fecha Vence: {col_fecha_vence}")
    print(f"  - Saldo      : {col_saldo}")

    recordatorios = []
    cliente_actual = correo_cliente_actual = None
    vendedor_actual = correo_vendedor_actual = "N/A"
    cupo_actual = 0

    sin_cliente = saldo_cero = 0
    vencidas = proximas = no_vencidas = 0
    hoy = date.today()

    print(f"\n[DEBUG] Clientes NO identificados en Excel Clientes:")
    print("-" * 80)

    for _, row in df.iterrows():
        cuenta_val = row[col_cuenta] if pd.notna(row[col_cuenta]) else None
        if cuenta_val is None:
            continue

        cuenta_str = str(cuenta_val).strip()

        # --- Fila de cabecera de cliente: contiene "NIT." ---
        if "NIT." in cuenta_str.upper():
            idx_nit = cuenta_str.upper().find("NIT.")
            nombre_original = cuenta_str[:idx_nit].rstrip(" -").strip()
            nit_part = cuenta_str[idx_nit + 4:]
            nit_raw = normalizar_nit(nit_part.split(" ")[0].split("-")[0])

            nombre_norm = normalizar_nombre(nombre_original)
            nit_key = f"__nit__{nit_raw}" if nit_raw else None

            cliente_info = None
            if nit_key and nit_key in dict_clientes:
                cliente_info = dict_clientes[nit_key]
            elif nombre_norm in dict_clientes:
                cliente_info = dict_clientes[nombre_norm]
            else:
                sin_cliente += 1
                print(f"  [{sin_cliente}] NO ENCONTRADO: '{nombre_original}' (NIT: {nit_raw})")

            if cliente_info:
                cliente_actual        = cliente_info["cliente"]
                correo_cliente_actual = cliente_info["correo_cliente"]
                vendedor_actual       = cliente_info.get("vendedor", "N/A")
                correo_vendedor_actual= cliente_info.get("correo_vendedor", "N/A")
                cupo_actual           = cliente_info.get("cupo", 0)
            else:
                cliente_actual = correo_cliente_actual = None
            continue

        # --- Fila de TOTAL: "TOTAL:" en la columna Fecha Vence ---
        fecha_vence_val = row[col_fecha_vence] if pd.notna(row[col_fecha_vence]) else None
        if fecha_vence_val and "total" in str(fecha_vence_val).lower():
            continue

        if not cliente_actual or not correo_cliente_actual:
            continue

        # --- Fila de movimiento ---
        documento = row[col_documento] if pd.notna(row[col_documento]) else None
        if not documento:
            continue

        saldo = row[col_saldo] if pd.notna(row[col_saldo]) else 0
        try:
            saldo_float = float(saldo)
            if saldo_float <= 0:
                saldo_cero += 1
                continue
        except:
            continue

        if not fecha_vence_val:
            continue

        try:
            vencimiento_date = pd.to_datetime(fecha_vence_val).date()
            dias = (vencimiento_date - hoy).days
        except:
            continue

        fecha_emision_val = row[col_fecha] if col_fecha and pd.notna(row[col_fecha]) else None
        try:
            emision_str = pd.to_datetime(fecha_emision_val).strftime("%d/%m/%Y") if fecha_emision_val else "N/A"
        except:
            emision_str = "N/A"

        saldo_formateado = f"${saldo_float:,.0f}"

        if dias < 0:
            estado = "vencido";    badge_class = "badge-danger";  vencidas += 1
        elif dias <= 5:
            estado = "proximo";    badge_class = "badge-warning"; proximas += 1
        else:
            estado = "no_vencido"; badge_class = "badge-success"; no_vencidas += 1

        recordatorios.append({
            "cliente":           cliente_actual,
            "correo_cliente":    correo_cliente_actual,
            "vendedor":          vendedor_actual,
            "correo_vendedor":   correo_vendedor_actual,
            "local":             "N/A",
            "numero_factura":    str(documento),
            "fecha_emision":     emision_str,
            "fecha_vencimiento": vencimiento_date.strftime("%d/%m/%Y"),
            "dias":              dias,
            "saldo":             saldo_formateado,
            "saldo_numerico":    saldo_float,
            "estado":            estado,
            "badge_class":       badge_class,
            "cupo":              cupo_actual
        })

    print("-" * 80)
    print(f"\n[INFO] Cartera (nuevo formato) procesada:")
    print(f"  - Total recordatorios: {len(recordatorios)}")
    print(f"    • Vencidas: {vencidas} | Próximas: {proximas} | No vencidas: {no_vencidas}")
    print(f"  - Sin cliente: {sin_cliente} | Saldo cero/negativo: {saldo_cero}")
    return recordatorios


def _leer_cartera_formato_antiguo(archivo_bytes, dict_clientes, dict_vendedores):
    """Lee el formato antiguo de cartera: hoja 'Cartera por edades'."""
    header_row = detectar_fila_header_cartera(archivo_bytes)
    print(f"[INFO] Leyendo Excel Cartera (formato antiguo) con header en fila {header_row}")
    df = pd.read_excel(BytesIO(archivo_bytes), sheet_name="Cartera por edades", header=header_row)

    col_nombre_tercero = buscar_columna_exacta(df, ["Nombre tercero", "Nombretercero", "Cliente"])
    col_numero_fac = buscar_columna_exacta(df, ["Numero FAC", "NumeroFAC", "Factura", "Numero Factura"])
    col_emision = buscar_columna_exacta(df, ["Emision", "Emisión", "Fecha Emision", "FechaEmision"])
    col_vencimiento = buscar_columna_exacta(df, ["Vencimiento", "Fecha Vencimiento", "FechaVencimiento"])
    col_saldo = buscar_columna_exacta(df, ["Saldo", "saldo"])
    col_vendedor = buscar_columna_exacta(df, ["Vendedor", "vendedor"])
    col_local = buscar_columna_exacta(df, ["Local", "local", "Sucursal", "sucursal"])

    columnas_faltantes = []
    if not col_nombre_tercero: columnas_faltantes.append("Nombre tercero")
    if not col_numero_fac: columnas_faltantes.append("Numero FAC")
    if not col_vencimiento: columnas_faltantes.append("Vencimiento")
    if not col_saldo: columnas_faltantes.append("Saldo")

    if columnas_faltantes:
        raise ValueError(f"Columnas faltantes: {', '.join(columnas_faltantes)}")

    print(f"[INFO] Columnas detectadas en Excel 2:")
    print(f"  - Nombre tercero: {col_nombre_tercero}")
    print(f"  - Numero FAC: {col_numero_fac}")
    print(f"  - Vencimiento: {col_vencimiento}")
    print(f"  - Saldo: {col_saldo}")

    recordatorios = []
    sin_cliente = 0
    vencimiento_vacio = 0
    saldo_cero = 0

    # Contadores por categoría
    vencidas = 0
    proximas = 0
    no_vencidas = 0

    hoy = date.today()
    print(f"\n[INFO] Fecha de HOY: {hoy.strftime('%d/%m/%Y')}")
    print(f"\n[DEBUG] Clientes NO identificados en Excel 1:")
    print("-" * 80)

    for _, row in df.iterrows():
        nombre_tercero = row[col_nombre_tercero] if pd.notna(row[col_nombre_tercero]) else None
        if not nombre_tercero:
            continue

        nombre_tercero_norm = normalizar_nombre(nombre_tercero)

        if nombre_tercero_norm not in dict_clientes:
            sin_cliente += 1
            print(f"  [{sin_cliente}] NO ENCONTRADO")
            print(f"       Original: '{nombre_tercero}'")
            print(f"       Normalizado: '{nombre_tercero_norm}'")
            print()
            continue

        cliente_info = dict_clientes[nombre_tercero_norm]
        correo_cliente = cliente_info["correo_cliente"]
        cliente_nombre = cliente_info["cliente"]
        cupo_cliente = cliente_info.get("cupo", 0)

        vendedor = row[col_vendedor] if col_vendedor and pd.notna(row[col_vendedor]) else None
        correo_vendedor = None

        if vendedor:
            vendedor_norm = normalizar_nombre(vendedor)
            if vendedor_norm in dict_vendedores:
                correo_vendedor = dict_vendedores[vendedor_norm]

        numero_fac = row[col_numero_fac] if pd.notna(row[col_numero_fac]) else "N/A"
        emision = row[col_emision] if col_emision and pd.notna(row[col_emision]) else None
        vencimiento = row[col_vencimiento] if pd.notna(row[col_vencimiento]) else None
        saldo = row[col_saldo] if pd.notna(row[col_saldo]) else 0

        if not pd.notna(vencimiento):
            vencimiento_vacio += 1
            continue

        try:
            saldo_float = float(saldo)
            if saldo_float == 0:
                saldo_cero += 1
                continue
        except:
            saldo_float = 0

        try:
            vencimiento_date = pd.to_datetime(vencimiento).date()
            dias = (vencimiento_date - hoy).days
        except Exception as e:
            print(f"[ERROR] Factura {numero_fac}: Error al calcular días: {e}")
            continue

        # CAMBIO: NO filtrar por ventana de días, procesar TODAS las facturas

        try:
            emision_str = pd.to_datetime(emision).strftime("%d/%m/%Y") if pd.notna(emision) else "N/A"
        except:
            emision_str = str(emision) if emision else "N/A"

        vencimiento_str = vencimiento_date.strftime("%d/%m/%Y")

        try:
            saldo_formateado = f"${saldo_float:,.0f}"
        except:
            saldo_formateado = "$0"

        # CAMBIO: Clasificar en 3 categorías
        if dias < 0:
            estado = "vencido"
            badge_class = "badge-danger"
            vencidas += 1
        elif dias <= 5:
            estado = "proximo"
            badge_class = "badge-warning"
            proximas += 1
        else:
            estado = "no_vencido"
            badge_class = "badge-success"
            no_vencidas += 1

        local = row[col_local] if col_local and pd.notna(row[col_local]) else "N/A"

        recordatorios.append({
            "cliente": cliente_nombre,
            "correo_cliente": correo_cliente,
            "vendedor": vendedor if vendedor else "N/A",
            "correo_vendedor": correo_vendedor if correo_vendedor else "N/A",
            "local": str(local),
            "numero_factura": str(numero_fac),
            "fecha_emision": emision_str,
            "fecha_vencimiento": vencimiento_str,
            "dias": dias,
            "saldo": saldo_formateado,
            "saldo_numerico": saldo_float,
            "estado": estado,
            "badge_class": badge_class,
            "cupo": cupo_cliente
        })

    print("-" * 80)

    print(f"\n[INFO] Excel 2 procesado:")
    print(f"  - Total recordatorios generados: {len(recordatorios)}")
    print(f"    • Vencidas (días < 0): {vencidas}")
    print(f"    • Próximas (0 <= días <= 5): {proximas}")
    print(f"    • No vencidas (días > 5): {no_vencidas}")
    print(f"  - Sin cliente (omitidos): {sin_cliente}")
    print(f"  - Vencimiento vacío: {vencimiento_vacio}")
    print(f"  - Saldo en cero: {saldo_cero}")

    return recordatorios


# ==========================================
# FUNCIONES DE ENVÍO DE CORREO
# ==========================================


def enviar_email_individual(destinatario_principal, lista_cc, asunto, cuerpo_html, cuerpo_texto=None):
    """Envía un correo electrónico individual con CC múltiple via Resend."""
    try:
        if not RESEND_API_KEY:
            return {
                "success": False,
                "destinatario": destinatario_principal,
                "error": "RESEND_API_KEY no configurada. Revisa el archivo .env"
            }

        if not destinatario_principal or "@" not in destinatario_principal:
            return {
                "success": False,
                "destinatario": destinatario_principal,
                "error": "Email de destinatario principal inválido"
            }

        params: resend.Emails.SendParams = {
            "from": f"{EMAIL_FROM_NAME} <{EMAIL_FROM_ADDRESS}>",
            "to": [destinatario_principal],
            "subject": asunto,
            "html": cuerpo_html,
        }

        cc_validos = [c for c in (lista_cc or []) if c and "@" in c]
        if cc_validos:
            params["cc"] = cc_validos

        if cuerpo_texto:
            params["text"] = cuerpo_texto

        resend.Emails.send(params)

        return {
            "success": True,
            "destinatario": destinatario_principal,
            "destinatario_cc": cc_validos,
            "error": None
        }

    except Exception as e:
        return {
            "success": False,
            "destinatario": destinatario_principal,
            "error": f"Error al enviar: {str(e)}"
        }


def generar_html_recordatorio_agrupado(cliente_agrupado, fecha_cartera=None, info_cierre=None, incluir_mensaje_cierre=True):
    """Genera HTML con TRES secciones: Vencidas, Próximas y No Vencidas."""
    cliente = cliente_agrupado.get("cliente", "Cliente")
    correo_vendedor = cliente_agrupado.get("correo_vendedor", "N/A")
    vendedor = cliente_agrupado.get("vendedor", "N/A")

    # Formatear fecha cartera para el título
    fecha_titulo = fecha_cartera if fecha_cartera else date.today().strftime("%d/%m/%Y")

    facturas_vencidas = cliente_agrupado.get("facturas_vencidas", [])
    facturas_proximas = cliente_agrupado.get("facturas_proximas", [])
    facturas_no_vencidas = cliente_agrupado.get("facturas_no_vencidas", [])

    total_facturas = cliente_agrupado.get("total_facturas", 0)
    total_vencidas = cliente_agrupado.get("total_vencidas", 0)
    total_proximas = cliente_agrupado.get("total_proximas", 0)
    total_no_vencidas = cliente_agrupado.get("total_no_vencidas", 0)
    total_saldo = cliente_agrupado.get("total_saldo", 0)
    cupo = cliente_agrupado.get("cupo", 0)
    cupo_disponible = cliente_agrupado.get("cupo_disponible", 0)

    logo_url = "https://images.jumpseller.com/store/lomarosa/store/logo/LR_LogotipoEslogan_CMYK.png?1662998750"

    # Generar mensaje de cierre trimestral si aplica
    mensaje_cierre_html = ""
    if info_cierre and info_cierre.get("es_cierre") and incluir_mensaje_cierre:
        tipo_cierre = info_cierre.get("mensaje_tipo", "Cierre trimestral")
        fecha_cierre = info_cierre.get("fecha_formateada", fecha_titulo)

        mensaje_cierre_html = f"""
        <div style="background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%); padding: 25px; margin: 20px 0 30px 0; border-radius: 12px; border-left: 5px solid #f59e0b;">
            <h3 style="color: #92400e; margin: 0 0 15px 0; font-size: 20px;">
                📋 {tipo_cierre} - {fecha_cierre}
            </h3>
            <p style="color: #78350f; margin: 0 0 12px 0; line-height: 1.7;">
                Queremos agradecerte por la confianza y el trabajo conjunto durante este periodo. Valoramos profundamente nuestra relación comercial y el crecimiento que hemos construido contigo.
            </p>
            <p style="color: #78350f; margin: 0 0 12px 0; line-height: 1.7;">
                Como parte de nuestros procesos de control y cierre contable, estamos realizando la confirmación de saldos con nuestros clientes y aliados estratégicos. Este ejercicio nos permite garantizar la precisión de nuestra información financiera y fortalecer la transparencia en nuestras relaciones comerciales.
            </p>
            <p style="color: #78350f; margin: 0; line-height: 1.7;">
                Agradecemos de antemano tu colaboración con esta verificación y quedamos atentos a cualquier inquietud que pueda surgir.
            </p>
        </div>
        """

    # Formatear montos
    total_saldo_formateado = f"${total_saldo:,.0f}"
    cupo_formateado = f"${cupo:,.0f}"
    cupo_disponible_formateado = f"${cupo_disponible:,.0f}"

    # Color para cupo disponible (rojo si es negativo, verde si es positivo)
    cupo_disponible_color = "#dc2626" if cupo_disponible < 0 else "#10b981"
    cupo_disponible_emoji = "⚠️" if cupo_disponible < 0 else "✅"

    def generar_tabla_facturas(facturas, titulo, color_bg, emoji):
        """Helper para generar tabla de facturas por categoría."""
        if len(facturas) == 0:
            return ""

        filas = ""
        for factura in facturas:
            filas += f"""
            <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="padding: 10px; font-weight: bold;">{factura['numero_factura']}</td>
                <td style="padding: 10px; text-align: center;">{factura['fecha_emision']}</td>
                <td style="padding: 10px; text-align: center;">{factura['fecha_vencimiento']}</td>
                <td style="padding: 10px; text-align: center;">{factura['dias']} días</td>
                <td style="padding: 10px; text-align: right; font-weight: bold;">{factura['saldo']}</td>
            </tr>
            """

        subtotal = sum(f["saldo_numerico"] for f in facturas)
        subtotal_formateado = f"${subtotal:,.0f}"

        return f"""
        <div style="margin: 30px 0;">
            <h3 style="color: {color_bg}; border-bottom: 3px solid {color_bg}; padding-bottom: 10px; margin-bottom: 15px;">
                {emoji} {titulo} ({len(facturas)})
            </h3>
            <table style="width: 100%; border-collapse: collapse; margin: 15px 0;">
                <thead>
                    <tr style="background-color: {color_bg}; color: white;">
                        <th style="padding: 12px; text-align: left;">Factura</th>
                        <th style="padding: 12px; text-align: center;">Emisión</th>
                        <th style="padding: 12px; text-align: center;">Vencimiento</th>
                        <th style="padding: 12px; text-align: center;">Días</th>
                        <th style="padding: 12px; text-align: right;">Saldo</th>
                    </tr>
                </thead>
                <tbody>
                    {filas}
                    <tr style="background-color: #f8f9fa; font-weight: bold; border-top: 2px solid {color_bg};">
                        <td colspan="4" style="text-align: right; padding: 12px;">SUBTOTAL:</td>
                        <td style="text-align: right; padding: 12px;">{subtotal_formateado}</td>
                    </tr>
                </tbody>
            </table>
        </div>
        """

    # Generar secciones solo si hay facturas
    seccion_vencidas = generar_tabla_facturas(
        facturas_vencidas,
        "FACTURAS VENCIDAS",
        "#dc2626",
        "🔴"
    )

    seccion_proximas = generar_tabla_facturas(
        facturas_proximas,
        "FACTURAS PRÓXIMAS A VENCER (≤ 5 días)",
        "#f59e0b",
        "🟡"
    )

    seccion_no_vencidas = generar_tabla_facturas(
        facturas_no_vencidas,
        "FACTURAS NO VENCIDAS (> 5 días)",
        "#10b981",
        "🟢"
    )

    total_saldo_formateado = f"${total_saldo:,.0f}"

    return f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            body {{font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 900px; margin: 0 auto; padding: 20px;}}
            .container {{background-color: white; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.1);}}
            .logo {{text-align: center; padding: 25px;}}
            .logo img {{max-width: 250px;}}
            .header {{background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; text-align: center; border-radius: 10px 10px 0 0;}}
            .header h1 {{margin: 0; font-size: 26px;}}
            .content {{padding: 30px;}}
            .resumen {{display: flex; justify-content: space-around; margin: 20px 0; background-color: #f8f9fa; padding: 20px; border-radius: 8px; flex-wrap: wrap;}}
            .resumen-item {{text-align: center; margin: 10px;}}
            .resumen-numero {{font-size: 32px; font-weight: bold; color: #667eea;}}
            .info-vendedor {{background-color: #e3f2fd; padding: 15px; margin: 20px 0; border-left: 4px solid #2196F3; border-radius: 4px;}}
            .footer {{background-color: #0f172a; color: #94a3b8; padding: 25px; text-align: center; border-radius: 0 0 10px 10px;}}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="logo">
                <img src="{logo_url}" alt="Lomarosa">
            </div>

            <div class="header">
                <h1>📧 Estado de Cuenta - {fecha_titulo}</h1>
                <p>Cliente: <strong>{cliente}</strong></p>
            </div>

            <div class="content">
                {mensaje_cierre_html}
                <p>Estimado Cliente <strong>{cliente}</strong>,</p>
                <p>A continuación presentamos el estado completo de sus facturas pendientes:</p>

                <div class="resumen">
                    <div class="resumen-item">
                        <div class="resumen-numero">{total_facturas}</div>
                        <div>Total Facturas</div>
                    </div>
                    <div class="resumen-item">
                        <div class="resumen-numero" style="color: #dc2626;">{total_vencidas}</div>
                        <div>🔴 Vencidas</div>
                    </div>
                    <div class="resumen-item">
                        <div class="resumen-numero" style="color: #f59e0b;">{total_proximas}</div>
                        <div>🟡 Próximas</div>
                    </div>
                    <div class="resumen-item">
                        <div class="resumen-numero" style="color: #10b981;">{total_no_vencidas}</div>
                        <div>🟢 No Vencidas</div>
                    </div>
                    <div class="resumen-item">
                        <div class="resumen-numero" style="color: #dc2626;">{total_saldo_formateado}</div>
                        <div>💰 Total Cartera</div>
                    </div>
                    <div class="resumen-item">
                        <div class="resumen-numero" style="color: {cupo_disponible_color};">{cupo_disponible_emoji} {cupo_disponible_formateado}</div>
                        <div>Cupo Disponible</div>
                    </div>
                </div>

                {seccion_vencidas}
                {seccion_proximas}
                {seccion_no_vencidas}

                <div style="background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%); padding: 20px; margin: 30px 0; border-radius: 8px; border-top: 4px solid #667eea;">
                    <h3 style="margin: 0 0 10px 0; text-align: center;">TOTAL GENERAL</h3>
                    <p style="font-size: 32px; font-weight: bold; text-align: center; margin: 0; color: #667eea;">{total_saldo_formateado}</p>
                    <p style="text-align: center; margin: 10px 0 0 0; font-size: 14px; color: #666;">Total de {total_facturas} facturas pendientes</p>
                </div>

                <div class="info-vendedor">
                    <strong>👤 Vendedor asignado:</strong> {vendedor}<br>
                    <strong>📧 Contacto:</strong> {correo_vendedor if correo_vendedor != 'N/A' else 'No asignado'}<br>
                    <strong>📞 Para consultas:</strong> Comuníquese con su vendedor<br>
                    <strong>⚠️ Dudas o solicitudes:</strong> Si cree que hay algo equivocado o quiere la cartera completa comuníquese con <a href="mailto:tesoreria@grupolom.com" style="color: #2196F3; text-decoration: none;">tesoreria@grupolom.com</a>
                </div>
            </div>

            <div class="footer">
                <p><strong>Lomarosa</strong><br>
                <em>Campo bien hecho, cerdos bien criados</em></p>
                <hr style="border: 1px solid #475569; margin: 15px 0;">
                <p style="font-size: 11px;">Este es un mensaje automático. No responder directamente a este correo.</p>
            </div>
        </div>
    </body>
    </html>
    """


def _enviar_lote_agrupado(recordatorios_agrupados, fecha_cartera=None, info_cierre=None, incluir_mensaje_cierre=True):
    """Envía lote de correos UNIFICADOS (vencidas + próximas + no vencidas)."""
    resultados = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        tareas = {}

        for cliente_agrupado in recordatorios_agrupados:
            destinatario_principal = cliente_agrupado.get("correo_cliente", "")
            correo_vendedor = cliente_agrupado.get("correo_vendedor", None)

            # CC siempre incluye cartera@grupolom.com + comercial del cliente
            lista_cc = [CC_CARTERA]
            if correo_vendedor and correo_vendedor != "N/A":
                lista_cc.append(correo_vendedor)

            total_facturas = cliente_agrupado.get("total_facturas", 0)
            total_vencidas = cliente_agrupado.get("total_vencidas", 0)
            total_proximas = cliente_agrupado.get("total_proximas", 0)

            fecha_asunto = fecha_cartera if fecha_cartera else date.today().strftime("%d/%m/%Y")
            asunto = f"Estado de Cuenta - {fecha_asunto} - {cliente_agrupado.get('cliente', 'Cliente')}"
            cuerpo_html = generar_html_recordatorio_agrupado(
                cliente_agrupado,
                fecha_cartera=fecha_cartera,
                info_cierre=info_cierre,
                incluir_mensaje_cierre=incluir_mensaje_cierre
            )
            cuerpo_texto = f"Tiene {total_facturas} facturas pendientes ({total_vencidas} vencidas, {total_proximas} próximas)"

            future = executor.submit(
                enviar_email_individual,
                destinatario_principal,
                lista_cc,
                asunto,
                cuerpo_html,
                cuerpo_texto
            )

            tareas[future] = cliente_agrupado

        for future in as_completed(tareas):
            cliente_agrupado = tareas[future]
            try:
                resultado = future.result()
                resultados.append({
                    "destinatario": resultado["destinatario"],
                    "cliente": cliente_agrupado.get("cliente"),
                    "facturas": cliente_agrupado.get("total_facturas", 0),
                    "vencidas": cliente_agrupado.get("total_vencidas", 0),
                    "proximas": cliente_agrupado.get("total_proximas", 0),
                    "no_vencidas": cliente_agrupado.get("total_no_vencidas", 0),
                    "success": resultado["success"],
                    "error": resultado["error"]
                })
            except Exception as e:
                resultados.append({
                    "destinatario": cliente_agrupado.get("correo_cliente"),
                    "cliente": cliente_agrupado.get("cliente"),
                    "facturas": cliente_agrupado.get("total_facturas", 0),
                    "vencidas": cliente_agrupado.get("total_vencidas", 0),
                    "proximas": cliente_agrupado.get("total_proximas", 0),
                    "no_vencidas": cliente_agrupado.get("total_no_vencidas", 0),
                    "success": False,
                    "error": str(e)
                })

    return resultados


# ==========================================
# RUTAS DE LA APLICACIÓN
# ==========================================


@app.route("/")
def index():
    """Renderiza la página principal."""
    return render_template("index.html")


@app.route("/test-email", methods=["GET"])
def test_email():
    """Prueba la configuración de Resend enviando un correo de prueba."""
    try:
        if not RESEND_API_KEY:
            return jsonify({
                "success": False,
                "message": "RESEND_API_KEY no configurada",
                "detalles": "Debes configurar RESEND_API_KEY en el archivo .env"
            }), 400

        destinatario_prueba = EMAIL_FROM_ADDRESS
        asunto = "Prueba de Configuración Resend - Cartera Lomarosa"

        cuerpo_html = """
        <html>
            <body style="font-family: Arial, sans-serif; padding: 20px;">
                <h2 style="color: #667eea;">✅ Configuración Resend Exitosa</h2>
                <p>Si estás leyendo este correo, significa que Resend está funcionando correctamente desde <strong>cartera@grupolom.com</strong>.</p>
                <hr>
                <p style="color: #666; font-size: 12px;">
                    Sistema de Recordatorios de Pago - Cartera Lomarosa
                </p>
            </body>
        </html>
        """

        resultado = enviar_email_individual(
            destinatario_principal=destinatario_prueba,
            lista_cc=[],
            asunto=asunto,
            cuerpo_html=cuerpo_html
        )

        if resultado["success"]:
            return jsonify({
                "success": True,
                "message": f"Correo de prueba enviado exitosamente a {destinatario_prueba}",
                "detalles": {
                    "remitente": f"{EMAIL_FROM_NAME} <{EMAIL_FROM_ADDRESS}>",
                    "destinatario": destinatario_prueba
                }
            })
        else:
            return jsonify({
                "success": False,
                "message": "Error al enviar correo de prueba",
                "error": resultado["error"]
            }), 500

    except Exception as e:
        return jsonify({
            "success": False,
            "message": "Error al probar configuración Resend",
            "error": str(e)
        }), 500


@app.route("/procesar-excel", methods=["POST"])
def procesar_excel():
    """Procesa ambos archivos Excel y retorna recordatorios con matching por nombre."""
    try:
        if 'file1' not in request.files or 'file2' not in request.files:
            return jsonify({
                "success": False,
                "message": "Faltan archivos. Debes enviar file1 y file2."
            }), 400
        
        file1 = request.files['file1']
        file2 = request.files['file2']
        
        contenido1 = file1.read()
        contenido2 = file2.read()

        # Detectar cuál archivo es cuál probando ambos
        archivo_clientes = None
        archivo_cartera = None

        # Variables para debug
        debug_log = []

        # Intentar detectar archivo 1
        try:
            # Primero probar si es archivo de clientes (sin hoja específica)
            df1_test = pd.read_excel(BytesIO(contenido1))
            tipo1_cliente = detectar_tipo_excel(df1_test, "(Archivo 1 como CLIENTES)")
            debug_log.append(f"Archivo1_clientes: {tipo1_cliente}, cols={list(df1_test.columns)[:5]}")
        except Exception as e:
            tipo1_cliente = None
            debug_log.append(f"Archivo1_clientes: ERROR - {str(e)}")

        try:
            sheet1_cartera = _detectar_sheet_cartera(contenido1)
            if sheet1_cartera:
                header_row_1 = detectar_fila_header_cartera(contenido1)
                df1_cartera = pd.read_excel(BytesIO(contenido1), sheet_name=sheet1_cartera, header=header_row_1)
                tipo1_cartera = detectar_tipo_excel(df1_cartera, f"(Archivo 1 como CARTERA, hoja='{sheet1_cartera}')")
                debug_log.append(f"Archivo1_cartera: {tipo1_cartera}, hoja={sheet1_cartera}, cols={list(df1_cartera.columns)[:5]}")
            else:
                tipo1_cartera = None
                debug_log.append("Archivo1_cartera: sin hoja de cartera reconocida")
        except Exception as e:
            tipo1_cartera = None
            debug_log.append(f"Archivo1_cartera: ERROR - {str(e)}")

        # Intentar detectar archivo 2
        try:
            df2_test = pd.read_excel(BytesIO(contenido2))
            tipo2_cliente = detectar_tipo_excel(df2_test, "(Archivo 2 como CLIENTES)")
            debug_log.append(f"Archivo2_clientes: {tipo2_cliente}, cols={list(df2_test.columns)[:5]}")
        except Exception as e:
            tipo2_cliente = None
            debug_log.append(f"Archivo2_clientes: ERROR - {str(e)}")

        try:
            sheet2_cartera = _detectar_sheet_cartera(contenido2)
            if sheet2_cartera:
                header_row_2 = detectar_fila_header_cartera(contenido2)
                df2_cartera = pd.read_excel(BytesIO(contenido2), sheet_name=sheet2_cartera, header=header_row_2)
                tipo2_cartera = detectar_tipo_excel(df2_cartera, f"(Archivo 2 como CARTERA, hoja='{sheet2_cartera}')")
                debug_log.append(f"Archivo2_cartera: {tipo2_cartera}, hoja={sheet2_cartera}, cols={list(df2_cartera.columns)[:5]}")
            else:
                tipo2_cartera = None
                debug_log.append("Archivo2_cartera: sin hoja de cartera reconocida")
        except Exception as e:
            tipo2_cartera = None
            debug_log.append(f"Archivo2_cartera: ERROR - {str(e)}")

        # Determinar qué archivo es cuál
        print(f"\n[RESUMEN DETECCIÓN]")
        print(f"  Archivo 1 - como clientes: {tipo1_cliente}")
        print(f"  Archivo 1 - como cartera: {tipo1_cartera}")
        print(f"  Archivo 2 - como clientes: {tipo2_cliente}")
        print(f"  Archivo 2 - como cartera: {tipo2_cartera}")

        if tipo1_cliente == "clientes" and tipo2_cartera == "cartera":
            archivo_clientes = contenido1
            archivo_cartera = contenido2
            print("[INFO] Archivo 1 = Clientes, Archivo 2 = Cartera")
        elif tipo1_cartera == "cartera" and tipo2_cliente == "clientes":
            archivo_clientes = contenido2
            archivo_cartera = contenido1
            print("[INFO] Archivo 1 = Cartera, Archivo 2 = Clientes")
        elif tipo2_cliente == "clientes" and tipo1_cartera == "cartera":
            archivo_clientes = contenido2
            archivo_cartera = contenido1
            print("[INFO] Archivo 1 = Cartera, Archivo 2 = Clientes")
        elif tipo2_cartera == "cartera" and tipo1_cliente == "clientes":
            archivo_clientes = contenido1
            archivo_cartera = contenido2
            print("[INFO] Archivo 1 = Clientes, Archivo 2 = Cartera")
        else:
            # Mensaje de error más detallado con debug completo
            debug_info = " | ".join(debug_log)
            return jsonify({
                "success": False,
                "message": f"No se pudieron detectar los tipos de archivo. DEBUG: {debug_info}"
            }), 400
        
        dict_clientes, dict_vendedores = leer_excel_clientes(archivo_clientes)
        recordatorios = leer_excel_cartera(archivo_cartera, dict_clientes, dict_vendedores)

        # Extraer fecha cartera y detectar cierre trimestral
        fecha_cartera = extraer_fecha_cartera(archivo_cartera)
        info_cierre = detectar_cierre_trimestral(fecha_cartera)

        print(f"\n[INFO] Fecha Cartera: {fecha_cartera.strftime('%d/%m/%Y') if fecha_cartera else 'No detectada'}")
        if info_cierre:
            print(f"[INFO] ¡CIERRE DETECTADO! Tipo: {info_cierre['mensaje_tipo']}")

        if not recordatorios:
            return jsonify({
                "success": True,
                "recordatorios": [],
                "stats": {
                    "total": 0,
                    "vencidas": 0,
                    "proximas": 0,
                    "no_vencidas": 0
                },
                "fecha_cartera": fecha_cartera.strftime("%d/%m/%Y") if fecha_cartera else None,
                "info_cierre": info_cierre,
                "message": "No se encontraron facturas con email asignado."
            })

        # Contar facturas por categoría
        vencidas = len([r for r in recordatorios if r["estado"] == "vencido"])
        proximas = len([r for r in recordatorios if r["estado"] == "proximo"])
        no_vencidas = len([r for r in recordatorios if r["estado"] == "no_vencido"])

        return jsonify({
            "success": True,
            "recordatorios": recordatorios,
            "stats": {
                "total": len(recordatorios),
                "vencidas": vencidas,
                "proximas": proximas,
                "no_vencidas": no_vencidas
            },
            "fecha_cartera": fecha_cartera.strftime("%d/%m/%Y") if fecha_cartera else None,
            "info_cierre": info_cierre
        })
    
    except Exception as e:
        print(f"[ERROR] Error al procesar Excel: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({
            "success": False,
            "message": "Error al procesar archivos Excel",
            "error": str(e)
        }), 500

@app.route("/enviar-correos", methods=["POST"])
def enviar_correos():
    """Envía correos UNIFICADOS por cliente (incluye vencidas + próximas + no vencidas)."""
    try:
        datos = request.get_json()

        if not datos or "recordatorios" not in datos:
            return jsonify({
                "success": False,
                "message": "Datos incorrectos"
            }), 400

        recordatorios = datos["recordatorios"]
        fecha_cartera = datos.get("fecha_cartera", None)
        info_cierre = datos.get("info_cierre", None)
        incluir_mensaje_cierre = datos.get("incluir_mensaje_cierre", True)

        print(f"\n[INFO] Parámetros de envío:")
        print(f"  - Fecha Cartera: {fecha_cartera}")
        print(f"  - Info Cierre: {info_cierre}")
        print(f"  - Incluir mensaje cierre: {incluir_mensaje_cierre}")

        if not isinstance(recordatorios, list) or len(recordatorios) == 0:
            return jsonify({
                "success": False,
                "message": "Lista vacía"
            }), 400

        if not RESEND_API_KEY:
            return jsonify({
                "success": False,
                "message": "RESEND_API_KEY no configurada en .env"
            }), 500

        # ← AGRUPAR por cliente + email (UN SOLO correo por cliente)
        print("\n[INFO] Agrupando recordatorios por cliente + email (unificado)...")
        recordatorios_agrupados = agrupar_recordatorios_por_cliente(recordatorios)

        print(f"\n[INFO] Iniciando envío de {len(recordatorios_agrupados)} correos unificados...")

        # ← ENVIAR LOTE
        resultados = _enviar_lote_agrupado(
            recordatorios_agrupados,
            fecha_cartera=fecha_cartera,
            info_cierre=info_cierre,
            incluir_mensaje_cierre=incluir_mensaje_cierre
        )

        exitosos = sum(1 for r in resultados if r["success"])
        fallidos = len(resultados) - exitosos

        return jsonify({
            "success": True,
            "message": f"✅ Envío completado: {len(recordatorios_agrupados)} correos unificados",
            "total": len(resultados),
            "exitosos": exitosos,
            "fallidos": fallidos,
            "resultados": resultados
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500




def abrir_navegador():
    """Abre el navegador en http://localhost:5000 después de 1.5 segundos."""
    webbrowser.open("http://localhost:5000")


if __name__ == "__main__":
    print("=" * 60)
    print("Sistema de Recordatorios de Pago - Cartera Lomarosa")
    print("=" * 60)
    print(f"Servidor iniciado en: http://localhost:5000")
    print(f"Proveedor de correo: Resend API")
    print(f"Remitente: {EMAIL_FROM_NAME} <{EMAIL_FROM_ADDRESS}>")
    print(f"API Key: {'[OK] Configurada' if RESEND_API_KEY else '[FALTA] NO CONFIGURADA (revisa .env)'}")
    print("=" * 60)
    print("\nPresiona Ctrl+C para detener el servidor.\n")
    
    Timer(1.5, abrir_navegador).start()
    
    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True,
        use_reloader=False
    )
