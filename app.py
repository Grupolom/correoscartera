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
# FUNCIONES DE AGRUPACIÓN
# ==========================================


def agrupar_recordatorios_por_cliente_y_estado(recordatorios):
    """Agrupa por cliente, email Y estado - Permite separar clientes con mismo email."""
    agrupados = {}

    for recordatorio in recordatorios:
        cliente_nombre = recordatorio.get("cliente")
        cliente_email = recordatorio.get("correo_cliente")
        estado = recordatorio.get("estado")

        # Key único: cliente + email + estado (permite separar clientes con mismo email)
        key = f"{cliente_nombre}|{cliente_email}|{estado}"

        if key not in agrupados:
            agrupados[key] = {
                "cliente": cliente_nombre,
                "correo_cliente": cliente_email,
                "vendedor": recordatorio.get("vendedor"),
                "correo_vendedor": recordatorio.get("correo_vendedor"),
                "local": recordatorio.get("local"),
                "estado": estado,
                "facturas": []
            }
        
        agrupados[key]["facturas"].append({
            "numero_factura": recordatorio.get("numero_factura"),
            "fecha_emision": recordatorio.get("fecha_emision"),
            "fecha_vencimiento": recordatorio.get("fecha_vencimiento"),
            "dias": recordatorio.get("dias"),
            "saldo": recordatorio.get("saldo"),
            "saldo_numerico": recordatorio.get("saldo_numerico"),
            "estado": recordatorio.get("estado")
        })
    
    resultado = list(agrupados.values())

    print(f"\n[INFO] Agrupación por cliente + email + estado:")
    print(f"  - Recordatorios individuales (facturas): {len(recordatorios)}")
    print(f"  - Correos a enviar (clientes separados): {len(resultado)}")
    vencidos = sum(1 for r in resultado if r['estado'] == 'vencido')
    proximos = sum(1 for r in resultado if r['estado'] == 'proximo')
    print(f"    • Vencidos: {vencidos}")
    print(f"    • Próximos: {proximos}")
    print(f"  - Nota: Clientes con mismo email se envían por separado")

    return resultado




def dividir_en_lotes(recordatorios, limite=450):
    """Divide los recordatorios en lotes de máximo 'limite' correos."""
    lote1 = recordatorios[:limite]
    lote2 = recordatorios[limite:]
    return lote1, lote2


# ==========================================
# FUNCIONES DE LECTURA DE EXCEL
# ==========================================


def detectar_tipo_excel(df):
    """Detecta si el Excel es Excel 1 (Clientes) o Excel 2 (Cartera) según sus columnas."""
    columnas_lower = [normalizar_columna(col) for col in df.columns]
    columnas_str = " ".join(columnas_lower)
    
    print("=" * 60)
    print(f"[DEBUG] Detectando tipo de Excel...")
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
    
    print(f"[DEBUG] Verificación Excel 1:")
    print(f"  - tiene_nit: {tiene_nit}")
    print(f"  - tiene_cliente: {tiene_cliente}")
    print(f"  - tiene_correo_cliente: {tiene_correo_cliente}")
    print()
    print(f"[DEBUG] Verificación Excel 2:")
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
    
    if not col_cliente:
        raise ValueError(f"No se encontró columna 'Cliente' en Excel 1. Columnas: {list(df.columns)}")
    if not col_correo_cliente:
        raise ValueError(f"No se encontró columna 'Correo cliente' en Excel 1. Columnas: {list(df.columns)}")
    
    print(f"[INFO] Columnas detectadas en Excel 1:")
    print(f"  - Cliente: {col_cliente}")
    print(f"  - Correo cliente: {col_correo_cliente}")
    print(f"  - Vendedor: {col_vendedor}")
    print(f"  - Correo vendedor: {col_correo_vendedor}")
    
    dict_clientes = {}
    dict_vendedores = {}
    
    for _, row in df.iterrows():
        cliente = row[col_cliente] if pd.notna(row[col_cliente]) else None
        correo_cliente = row[col_correo_cliente] if pd.notna(row[col_correo_cliente]) else None
        
        if cliente and correo_cliente:
            cliente_norm = normalizar_nombre(cliente)
            if cliente_norm:
                dict_clientes[cliente_norm] = {
                    "nit": str(row[col_nit]).strip() if col_nit and pd.notna(row[col_nit]) else "N/A",
                    "cliente": str(cliente).strip(),
                    "nombre_comercial": str(row[col_nombre_comercial]).strip() if col_nombre_comercial and pd.notna(row[col_nombre_comercial]) else "N/A",
                    "correo_cliente": str(correo_cliente).strip(),
                    "canal": str(row[col_canal]).strip() if col_canal and pd.notna(row[col_canal]) else "N/A"
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
    """Lee Excel 2 (Cartera) - Calcula días desde FECHAS REALES, NO desde columna Días."""
    df = pd.read_excel(BytesIO(archivo_bytes), sheet_name="Cartera por edades", header=11)
    
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
    fuera_ventana = 0
    vencimiento_vacio = 0
    saldo_cero = 0
    
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
        
        if dias > 5:
            fuera_ventana += 1
            continue
        
        try:
            emision_str = pd.to_datetime(emision).strftime("%d/%m/%Y") if pd.notna(emision) else "N/A"
        except:
            emision_str = str(emision) if emision else "N/A"
        
        vencimiento_str = vencimiento_date.strftime("%d/%m/%Y")
        
        try:
            saldo_formateado = f"${saldo_float:,.0f}"
        except:
            saldo_formateado = "$0"
        
        if dias < 0:
            estado = "vencido"
            badge_class = "badge-danger"
        else:
            estado = "proximo"
            badge_class = "badge-warning"
        
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
            "badge_class": badge_class
        })
    
    print("-" * 80)
    
    vencidos = len([r for r in recordatorios if r["estado"] == "vencido"])
    proximos = len([r for r in recordatorios if r["estado"] == "proximo"])
    
    print(f"\n[INFO] Excel 2 procesado:")
    print(f"  - Recordatorios generados: {len(recordatorios)}")
    print(f"    • Vencidos (días < 0): {vencidos}")
    print(f"    • Próximos (0 <= días <= 5): {proximos}")
    print(f"  - Sin cliente (omitidos): {sin_cliente}")
    print(f"  - Vencimiento vacío: {vencimiento_vacio}")
    print(f"  - Saldo en cero: {saldo_cero}")
    print(f"  - Fuera de ventana (>5 días): {fuera_ventana}")
    
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


def generar_html_recordatorio_agrupado(cliente_agrupado):
    """Genera HTML con MÚLTIPLES facturas en UN SOLO correo."""
    cliente = cliente_agrupado.get("cliente", "Cliente")
    correo_vendedor = cliente_agrupado.get("correo_vendedor", "N/A")
    vendedor = cliente_agrupado.get("vendedor", "N/A")
    facturas = cliente_agrupado.get("facturas", [])
    
    logo_url = "https://images.jumpseller.com/store/lomarosa/store/logo/LR_LogotipoEslogan_CMYK.png?1662998750"
    
    limite_tabla = 50
    facturas_mostradas = facturas[:limite_tabla]
    facturas_ocultas = len(facturas) - limite_tabla
    
    filas_facturas = ""
    
    for factura in facturas_mostradas:
        estado_emoji = "🔴" if factura["estado"] == "vencido" else "🟡"
        estado_texto = "VENCIDO" if factura["estado"] == "vencido" else "PRÓXIMO"
        
        filas_facturas += f"""
        <tr style="border-bottom: 1px solid #e0e0e0;">
            <td style="padding: 10px; text-align: center;">{estado_emoji} {estado_texto}</td>
            <td style="padding: 10px; font-weight: bold;">{factura['numero_factura']}</td>
            <td style="padding: 10px; text-align: center;">{factura['fecha_emision']}</td>
            <td style="padding: 10px; text-align: center;">{factura['fecha_vencimiento']}</td>
            <td style="padding: 10px; text-align: center;">{factura['dias']} días</td>
            <td style="padding: 10px; text-align: right; color: #dc2626; font-weight: bold;">{factura['saldo']}</td>
        </tr>
        """
    
    total_saldo_real = sum(f["saldo_numerico"] for f in facturas)
    total_saldo_formateado = f"${total_saldo_real:,.0f}"
    
    advertencia_ocultas = ""
    if facturas_ocultas > 0:
        advertencia_ocultas = f"""
        <div style="background-color: #fff3cd; padding: 15px; margin: 15px 0; border-left: 4px solid #ffc107; border-radius: 4px;">
            <strong>ℹ️ Información:</strong> Este correo muestra las primeras {limite_tabla} facturas de {len(facturas)} totales.
            Las restantes {facturas_ocultas} facturas están incluidas en el total.
        </div>
        """
    
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
            .tabla-facturas {{width: 100%; border-collapse: collapse; margin: 20px 0;}}
            .tabla-facturas th {{background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 12px; text-align: left; font-weight: 600;}}
            .tabla-facturas td {{padding: 10px 12px; font-size: 14px;}}
            .total-row {{background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%); font-weight: bold; border-top: 3px solid #667eea;}}
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
                <h1>📧 Recordatorio de Vencimiento de Facturas</h1>
                <p>Cliente: <strong>{cliente}</strong></p>
            </div>
            
            <div class="content">
                <p>Estimado Cliente <strong>{cliente}</strong>,</p>
                <p>A continuación presentamos el detalle de sus facturas vencidas y próximas a vencer:</p>
                
                <div class="resumen">
                    <div class="resumen-item">
                        <div class="resumen-numero">{len(facturas)}</div>
                        <div>Facturas Totales</div>
                    </div>
                    <div class="resumen-item">
                        <div class="resumen-numero" style="color: #dc2626;">{sum(1 for f in facturas if f['estado'] == 'vencido')}</div>
                        <div>Vencidas</div>
                    </div>
                    <div class="resumen-item">
                        <div class="resumen-numero" style="color: #f59e0b;">{sum(1 for f in facturas if f['estado'] == 'proximo')}</div>
                        <div>Próximas</div>
                    </div>
                    <div class="resumen-item">
                        <div class="resumen-numero" style="color: #dc2626;">{total_saldo_formateado}</div>
                        <div>Saldo Total</div>
                    </div>
                </div>
                
                {advertencia_ocultas}
                
                <table class="tabla-facturas">
                    <thead>
                        <tr>
                            <th>Estado</th>
                            <th>Factura</th>
                            <th>Emisión</th>
                            <th>Vencimiento</th>
                            <th>Días</th>
                            <th style="text-align: right;">Saldo</th>
                        </tr>
                    </thead>
                    <tbody>
                        {filas_facturas}
                        <tr class="total-row">
                            <td colspan="5" style="text-align: right; padding: 15px;"><strong>TOTAL GENERAL:</strong></td>
                            <td style="text-align: right; padding: 15px; font-size: 18px;"><strong>{total_saldo_formateado}</strong></td>
                        </tr>
                    </tbody>
                </table>
                
                <div class="info-vendedor">
                    <strong>👤 Vendedor asignado:</strong> {vendedor}<br>
                    <strong>📧 Contacto:</strong> {correo_vendedor if correo_vendedor != 'N/A' else 'No asignado'}<br>
                    <strong>📞 Para consultas:</strong> Comuníquese con su vendedor
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


def _enviar_lote_agrupado(recordatorios_agrupados):
    """Envía lote de correos AGRUPADOS."""
    resultados = []
    
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        tareas = {}
        
        for cliente_agrupado in recordatorios_agrupados:
            destinatario_principal = cliente_agrupado.get("correo_cliente", "")
            destinatario_cc = cliente_agrupado.get("correo_vendedor", None)
            
            asunto = f"Recordatorio de Vencimiento - {len(cliente_agrupado['facturas'])} facturas - {cliente_agrupado.get('cliente', 'Cliente')}"
            cuerpo_html = generar_html_recordatorio_agrupado(cliente_agrupado)
            cuerpo_texto = f"Tiene {len(cliente_agrupado['facturas'])} facturas vencidas o próximas a vencer"
            
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
                    "facturas": len(cliente_agrupado.get("facturas", [])),
                    "success": resultado["success"],
                    "error": resultado["error"]
                })
            except Exception as e:
                resultados.append({
                    "destinatario": cliente_agrupado.get("correo_cliente"),
                    "cliente": cliente_agrupado.get("cliente"),
                    "facturas": len(cliente_agrupado.get("facturas", [])),
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
        
        df1 = pd.read_excel(BytesIO(contenido1))
        
        try:
            df2 = pd.read_excel(BytesIO(contenido2), sheet_name="Cartera por edades", header=11)
            print("[INFO] Excel 2: Leyendo hoja 'Cartera por edades' (desde fila 12) ✓")
        except Exception as e:
            return jsonify({
                "success": False,
                "message": f"No se encontró la hoja 'Cartera por edades' en el Excel 2. Error: {str(e)}"
            }), 400
        
        tipo1 = detectar_tipo_excel(df1)
        tipo2 = detectar_tipo_excel(df2)
        
        print(f"[INFO] Archivo 1 detectado como: {tipo1}")
        print(f"[INFO] Archivo 2 detectado como: {tipo2}")
        
        if tipo1 == "clientes" and tipo2 == "cartera":
            archivo_clientes = contenido1
            archivo_cartera = contenido2
        elif tipo1 == "cartera" and tipo2 == "clientes":
            archivo_clientes = contenido2
            archivo_cartera = contenido1
        else:
            return jsonify({
                "success": False,
                "message": f"No se pudieron detectar los tipos de archivo correctamente. Tipo1: {tipo1}, Tipo2: {tipo2}."
            }), 400
        
        dict_clientes, dict_vendedores = leer_excel_clientes(archivo_clientes)
        recordatorios = leer_excel_cartera(archivo_cartera, dict_clientes, dict_vendedores)
        
        if not recordatorios:
            return jsonify({
                "success": True,
                "recordatorios": [],
                "stats": {
                    "total": 0,
                    "vencidos": 0,
                    "proximos": 0
                },
                "message": "No se encontraron facturas próximas a vencer o vencidas con email asignado."
            })
        
        vencidos = len([r for r in recordatorios if r["estado"] == "vencido"])
        proximos = len([r for r in recordatorios if r["estado"] == "proximo"])
        
        return jsonify({
            "success": True,
            "recordatorios": recordatorios,
            "stats": {
                "total": len(recordatorios),
                "vencidos": vencidos,
                "proximos": proximos
            }
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
    """Envía correos AGRUPADOS por cliente Y estado (vencido/próximo)."""
    try:
        datos = request.get_json()
        
        if not datos or "recordatorios" not in datos:
            return jsonify({
                "success": False,
                "message": "Datos incorrecto"
            }), 400
        
        recordatorios = datos["recordatorios"]
        
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
        
        # ← AGRUPAR por cliente + email + estado (separa clientes con mismo email)
        print("\n[INFO] Agrupando recordatorios por cliente + email + estado...")
        recordatorios_agrupados = agrupar_recordatorios_por_cliente_y_estado(recordatorios)
        
        print(f"\n[INFO] Iniciando envío de {len(recordatorios_agrupados)} correos agrupados...")
        
        # ← ENVIAR LOTE
        resultados = _enviar_lote_agrupado(recordatorios_agrupados)
        
        exitosos = sum(1 for r in resultados if r["success"])
        fallidos = len(resultados) - exitosos
        
        return jsonify({
            "success": True,
            "message": f"✅ Envío completado: {len(recordatorios_agrupados)} correos consolidados",
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
