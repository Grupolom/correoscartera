import os
import json
import smtplib
import webbrowser
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
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
EMAIL_HOST = os.getenv("EMAIL_HOST", "smtp.gmail.com")
EMAIL_PORT = int(os.getenv("EMAIL_PORT", "587"))
EMAIL_USER = os.getenv("EMAIL_USER", "")  
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "")  
EMAIL_FROM_NAME = os.getenv("EMAIL_FROM_NAME", "Cartera Lomarosa")
EMAIL_FROM_ADDRESS = os.getenv("EMAIL_FROM_ADDRESS", EMAIL_USER)


MAX_WORKERS = int(os.getenv("MAX_WORKERS", "3"))


# ==========================================
# FUNCIONES DE NORMALIZACIÓN
# ==========================================


def normalizar_nombre(nombre):
    """Normaliza un nombre para hacer matching: trim + lowercase"""
    if not nombre:
        return ""
    return str(nombre).strip().lower()


def normalizar_columna(col):
    """Normaliza nombre de columna para búsqueda flexible"""
    return str(col).strip().lower().replace('  ', ' ')


# ==========================================
# FUNCIONES DE FECHA CARTERA Y CIERRE TRIMESTRAL
# ==========================================


def extraer_fecha_cartera(archivo_bytes):
    """
    Extrae la fecha de cartera del Excel buscando 'Fecha Cartera' en columna E
    y tomando el valor de la celda a su derecha (columna F).
    """
    try:
        # Leer el Excel sin header para buscar en todas las filas
        df_raw = pd.read_excel(BytesIO(archivo_bytes), sheet_name="Cartera por edades", header=None)

        # Buscar "Fecha Cartera" en columna E (índice 4)
        fecha_cartera = None
        for idx, valor in enumerate(df_raw.iloc[:, 4]):  # Columna E = índice 4
            if pd.notna(valor) and "fecha cartera" in str(valor).lower():
                # Tomar el valor de la celda a la derecha (columna F = índice 5)
                fecha_cartera_raw = df_raw.iloc[idx, 5]
                if pd.notna(fecha_cartera_raw):
                    try:
                        fecha_cartera = pd.to_datetime(fecha_cartera_raw).date()
                        print(f"[INFO] Fecha Cartera encontrada: {fecha_cartera.strftime('%d/%m/%Y')}")
                    except:
                        fecha_cartera = None
                        print(f"[WARNING] No se pudo parsear fecha cartera: {fecha_cartera_raw}")
                break

        if not fecha_cartera:
            print("[WARNING] No se encontró 'Fecha Cartera' en columna E del Excel")
            fecha_cartera = date.today()

        return fecha_cartera

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


def detectar_fila_header_cartera(archivo_bytes):
    """
    Detecta dinámicamente en qué fila está el header de la tabla de cartera.
    Busca la fila que contenga 'Nombre tercero' en la primera columna.
    Retorna el índice de la fila (para usar como header en pd.read_excel).
    """
    try:
        df_raw = pd.read_excel(BytesIO(archivo_bytes), sheet_name="Cartera por edades", header=None, nrows=20)

        for idx, valor in enumerate(df_raw.iloc[:, 0]):  # Columna A
            if pd.notna(valor) and "nombre tercero" in str(valor).lower():
                print(f"[DEBUG] Header de cartera encontrado en fila {idx}")
                return idx

        print("[DEBUG] No se encontró 'Nombre tercero' en columna A, usando default 11")
        return 11  # Default antiguo
    except Exception as e:
        print(f"[DEBUG] Error buscando header: {e}")
        return 11


def detectar_tipo_excel(df, debug_info=""):
    """Detecta si el Excel es Excel 1 (Clientes) o Excel 2 (Cartera) según sus columnas."""
    columnas_lower = [normalizar_columna(col) for col in df.columns]
    columnas_str = " ".join(columnas_lower)

    print("=" * 60)
    print(f"[DEBUG] Detectando tipo de Excel {debug_info}...")
    print(f"[DEBUG] Total columnas: {len(columnas_lower)}")
    print(f"[DEBUG] Primeras 15 columnas: {columnas_lower[:15]}")
    print("=" * 60)

    tiene_nit = "nit" in columnas_str
    tiene_cliente = "cliente" in columnas_str
    tiene_correo_cliente = "correo cliente" in columnas_str or "correocliente" in columnas_str.replace(' ', '')

    tiene_nombre_tercero = "nombre tercero" in columnas_str or "nombretercero" in columnas_str.replace(' ', '')
    tiene_numero_fac = "numero fac" in columnas_str or "numerofac" in columnas_str.replace(' ', '') or " fac " in columnas_str
    tiene_vencimiento = "vencimiento" in columnas_str
    tiene_dias = "dias" in columnas_str or "días" in columnas_str
    tiene_saldo = "saldo" in columnas_str

    print(f"[DEBUG] Verificación Excel CLIENTES:")
    print(f"  - tiene_nit: {tiene_nit}")
    print(f"  - tiene_cliente: {tiene_cliente}")
    print(f"  - tiene_correo_cliente: {tiene_correo_cliente}")
    print()
    print(f"[DEBUG] Verificación Excel CARTERA:")
    print(f"  - tiene_nombre_tercero: {tiene_nombre_tercero}")
    print(f"  - tiene_numero_fac: {tiene_numero_fac}")
    print(f"  - tiene_vencimiento: {tiene_vencimiento}")
    print(f"  - tiene_dias: {tiene_dias}")
    print(f"  - tiene_saldo: {tiene_saldo}")
    print("=" * 60)

    if tiene_nit and tiene_cliente and tiene_correo_cliente:
        print("[DEBUG] ✓ Detectado como: CLIENTES")
        return "clientes"
    elif tiene_nombre_tercero and tiene_numero_fac and tiene_vencimiento and tiene_dias and tiene_saldo:
        print("[DEBUG] ✓ Detectado como: CARTERA")
        return "cartera"
    else:
        print("[DEBUG] ✗ NO DETECTADO (devolviendo None)")
        print(f"[DEBUG] Columnas faltantes para CARTERA:")
        if not tiene_nombre_tercero: print("  - FALTA: nombre tercero")
        if not tiene_numero_fac: print("  - FALTA: numero fac")
        if not tiene_vencimiento: print("  - FALTA: vencimiento")
        if not tiene_dias: print("  - FALTA: dias")
        if not tiene_saldo: print("  - FALTA: saldo")
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
    """Lee Excel 1 (Clientes y Vendedores) y retorna dos diccionarios."""
    df = pd.read_excel(BytesIO(archivo_bytes))

    print(f"[DEBUG] Columnas en Excel 1: {list(df.columns)}")

    col_nit = buscar_columna_exacta(df, ["Nit", "NIT"])
    col_cliente = buscar_columna_exacta(df, ["Cliente", "cliente"])
    col_nombre_comercial = buscar_columna_exacta(df, ["Nombre comercial", "Nombrecomercial"])
    col_correo_cliente = buscar_columna_exacta(df, ["Correo cliente", "Correocliente", "Email cliente"])
    col_vendedor = buscar_columna_exacta(df, ["Vendedor", "vendedor"])
    col_correo_vendedor = buscar_columna_exacta(df, ["Correo vendedor", "Correovendedor", "Email vendedor"])
    col_canal = buscar_columna_exacta(df, ["Canal", "canal"])
    col_cupo = buscar_columna_exacta(df, ["Cupo", "cupo", "Cupo de crédito", "Cupo de credito", "Cupo credito"])

    if not col_cliente:
        raise ValueError(f"No se encontró columna 'Cliente' en Excel 1. Columnas: {list(df.columns)}")
    if not col_correo_cliente:
        raise ValueError(f"No se encontró columna 'Correo cliente' en Excel 1. Columnas: {list(df.columns)}")

    print(f"[INFO] Columnas detectadas en Excel 1:")
    print(f"  - Cliente: {col_cliente}")
    print(f"  - Correo cliente: {col_correo_cliente}")
    print(f"  - Vendedor: {col_vendedor}")
    print(f"  - Correo vendedor: {col_correo_vendedor}")
    print(f"  - Cupo: {col_cupo if col_cupo else '❌ NO ENCONTRADO (se usará $0)'}")

    dict_clientes = {}
    dict_vendedores = {}

    for _, row in df.iterrows():
        cliente = row[col_cliente] if pd.notna(row[col_cliente]) else None
        correo_cliente = row[col_correo_cliente] if pd.notna(row[col_correo_cliente]) else None

        if cliente and correo_cliente:
            cliente_norm = normalizar_nombre(cliente)
            if cliente_norm:
                # Procesar cupo (validar que sea numérico)
                cupo_valor = 0
                if col_cupo and pd.notna(row[col_cupo]):
                    try:
                        cupo_valor = float(row[col_cupo])
                    except (ValueError, TypeError):
                        cupo_valor = 0
                        print(f"[WARNING] Cupo inválido para cliente '{cliente}': {row[col_cupo]}")

                dict_clientes[cliente_norm] = {
                    "nit": str(row[col_nit]).strip() if col_nit and pd.notna(row[col_nit]) else "N/A",
                    "cliente": str(cliente).strip(),
                    "nombre_comercial": str(row[col_nombre_comercial]).strip() if col_nombre_comercial and pd.notna(row[col_nombre_comercial]) else "N/A",
                    "correo_cliente": str(correo_cliente).strip(),
                    "canal": str(row[col_canal]).strip() if col_canal and pd.notna(row[col_canal]) else "N/A",
                    "cupo": cupo_valor
                }

        if col_vendedor and col_correo_vendedor:
            vendedor = row[col_vendedor] if pd.notna(row[col_vendedor]) else None
            correo_vendedor = row[col_correo_vendedor] if pd.notna(row[col_correo_vendedor]) else None

            if vendedor and correo_vendedor:
                vendedor_norm = normalizar_nombre(vendedor)
                if vendedor_norm:
                    dict_vendedores[vendedor_norm] = str(correo_vendedor).strip()

    print(f"[INFO] Excel 1 procesado: {len(dict_clientes)} clientes, {len(dict_vendedores)} vendedores")

    return dict_clientes, dict_vendedores


def leer_excel_cartera(archivo_bytes, dict_clientes, dict_vendedores):
    """Lee Excel 2 (Cartera) - Procesa TODAS las facturas (vencidas, próximas y no vencidas)."""
    # Detectar dinámicamente la fila del header
    header_row = detectar_fila_header_cartera(archivo_bytes)
    print(f"[INFO] Leyendo Excel Cartera con header en fila {header_row}")
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


def crear_mensaje_email(destinatario_principal, destinatario_cc, asunto, cuerpo_html, cuerpo_texto=None):
    """Crea un mensaje de email en formato MIME con CC."""
    mensaje = MIMEMultipart("alternative")
    mensaje["Subject"] = asunto
    mensaje["From"] = f"{EMAIL_FROM_NAME} <{EMAIL_FROM_ADDRESS}>"
    mensaje["To"] = destinatario_principal
    
    if destinatario_cc:
        mensaje["Cc"] = destinatario_cc
    
    if cuerpo_texto:
        parte_texto = MIMEText(cuerpo_texto, "plain", "utf-8")
        mensaje.attach(parte_texto)
    
    parte_html = MIMEText(cuerpo_html, "html", "utf-8")
    mensaje.attach(parte_html)
    
    return mensaje


def enviar_email_individual(destinatario_principal, destinatario_cc, asunto, cuerpo_html, cuerpo_texto=None):
    """Envía un correo electrónico individual con CC opcional."""
    try:
        if not EMAIL_USER or not EMAIL_PASSWORD:
            return {
                "success": False,
                "destinatario": destinatario_principal,
                "error": "Credenciales de correo no configuradas. Revisa el archivo .env"
            }
        
        if not destinatario_principal or "@" not in destinatario_principal:
            return {
                "success": False,
                "destinatario": destinatario_principal,
                "error": "Email de destinatario principal inválido"
            }
        
        mensaje = crear_mensaje_email(destinatario_principal, destinatario_cc, asunto, cuerpo_html, cuerpo_texto)
        
        destinatarios = [destinatario_principal]
        if destinatario_cc and "@" in destinatario_cc:
            destinatarios.append(destinatario_cc)
        
        with smtplib.SMTP(EMAIL_HOST, EMAIL_PORT, timeout=30) as server:
            server.ehlo()
            server.starttls()  
            server.ehlo()
            server.login(EMAIL_USER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_FROM_ADDRESS, destinatarios, mensaje.as_string())
        
        return {
            "success": True,
            "destinatario": destinatario_principal,
            "destinatario_cc": destinatario_cc,
            "error": None
        }
    
    except smtplib.SMTPAuthenticationError:
        return {
            "success": False,
            "destinatario": destinatario_principal,
            "error": "Error de autenticación SMTP. Verifica tu correo y contraseña de aplicación."
        }
    
    except smtplib.SMTPException as e:
        return {
            "success": False,
            "destinatario": destinatario_principal,
            "error": f"Error SMTP: {str(e)}"
        }
    
    except Exception as e:
        return {
            "success": False,
            "destinatario": destinatario_principal,
            "error": f"Error inesperado: {str(e)}"
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
            destinatario_cc = cliente_agrupado.get("correo_vendedor", None)

            total_facturas = cliente_agrupado.get("total_facturas", 0)
            total_vencidas = cliente_agrupado.get("total_vencidas", 0)
            total_proximas = cliente_agrupado.get("total_proximas", 0)

            # Generar asunto descriptivo con fecha cartera
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
                destinatario_cc,
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
    """Prueba la configuración SMTP enviando un correo de prueba."""
    try:
        if not EMAIL_USER or not EMAIL_PASSWORD:
            return jsonify({
                "success": False,
                "message": "Credenciales de correo no configuradas",
                "detalles": "Debes configurar EMAIL_USER y EMAIL_PASSWORD en el archivo .env"
            }), 400
        
        email_prueba = EMAIL_USER
        asunto = "Prueba de Configuración SMTP - Cartera Lomarosa"
        
        cuerpo_html = """
        <html>
            <body style="font-family: Arial, sans-serif; padding: 20px;">
                <h2 style="color: #667eea;">✅ Configuración SMTP Exitosa</h2>
                <p>Si estás leyendo este correo, significa que tu configuración SMTP está funcionando correctamente.</p>
                <hr>
                <p style="color: #666; font-size: 12px;">
                    Sistema de Recordatorios de Pago - Cartera Lomarosa
                </p>
            </body>
        </html>
        """
        
        cuerpo_texto = "✅ Configuración SMTP Exitosa\n\nSi estás leyendo este correo, significa que tu configuración SMTP está funcionando correctamente."
        
        resultado = enviar_email_individual(
            destinatario_principal=email_prueba,
            destinatario_cc=None,
            asunto=asunto,
            cuerpo_html=cuerpo_html,
            cuerpo_texto=cuerpo_texto
        )
        
        if resultado["success"]:
            return jsonify({
                "success": True,
                "message": f"Correo de prueba enviado exitosamente a {email_prueba}",
                "detalles": {
                    "servidor": EMAIL_HOST,
                    "puerto": EMAIL_PORT,
                    "usuario": EMAIL_USER,
                    "destinatario": email_prueba
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
            "message": "Error al probar configuración SMTP",
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
            # Probar si es archivo de cartera (con hoja específica y header dinámico)
            header_row_1 = detectar_fila_header_cartera(contenido1)
            df1_cartera = pd.read_excel(BytesIO(contenido1), sheet_name="Cartera por edades", header=header_row_1)
            tipo1_cartera = detectar_tipo_excel(df1_cartera, f"(Archivo 1 como CARTERA, header={header_row_1})")
            debug_log.append(f"Archivo1_cartera: {tipo1_cartera}, header={header_row_1}, cols={list(df1_cartera.columns)[:5]}")
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
            # Probar si es archivo de cartera (con hoja específica y header dinámico)
            header_row_2 = detectar_fila_header_cartera(contenido2)
            df2_cartera = pd.read_excel(BytesIO(contenido2), sheet_name="Cartera por edades", header=header_row_2)
            tipo2_cartera = detectar_tipo_excel(df2_cartera, f"(Archivo 2 como CARTERA, header={header_row_2})")
            debug_log.append(f"Archivo2_cartera: {tipo2_cartera}, header={header_row_2}, cols={list(df2_cartera.columns)[:5]}")
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

        if not EMAIL_USER or not EMAIL_PASSWORD:
            return jsonify({
                "success": False,
                "message": "Credenciales no configuradas"
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
    print(f"Configuración SMTP: {EMAIL_HOST}:{EMAIL_PORT}")
    print(f"Usuario de correo: {EMAIL_USER if EMAIL_USER else '❌ NO CONFIGURADO'}")
    print("=" * 60)
    print("\nPresiona Ctrl+C para detener el servidor.\n")
    
    Timer(1.5, abrir_navegador).start()
    
    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True,
        use_reloader=False
    )
