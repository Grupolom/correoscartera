# Sistema de Recordatorios de Pago - Cartera Lomarosa

Sistema web automatizado para generar y enviar recordatorios de pago a terceros basándose en la cartera por edades y fechas de vencimiento. **NO requiere Power Automate ni descarga de CSV** - todo el envío de correos se hace directamente desde la aplicación.

## Características

- ✉️ **Envío automático de correos** con plantillas HTML profesionales
- 🎯 **Drag & Drop** intuitivo para cargar archivos Excel
- 📊 **Vista previa** con estadísticas de facturas próximas a vencer y vencidas
- 🔄 **Envío en paralelo** de correos electrónicos
- 📧 **Compatible con Gmail y Outlook** (Office 365)
- 🔒 **Procesamiento 100% local** - nada se sube a internet
- 📈 **Reportes detallados** de envíos exitosos y fallidos
- 🚀 **Un solo click** para instalar y ejecutar (`iniciar.bat`)

## Requisitos

- **Python 3.8+** (descarga desde [python.org](https://www.python.org/downloads/))
- **Navegador moderno** (Chrome, Firefox, Edge)
- **Cuenta de correo** Gmail u Outlook

## Instalación Rápida

### Método 1: Automático (Recomendado)

1. Descarga o clona este proyecto
2. Haz **doble clic** en `iniciar.bat`
3. El script hará todo automáticamente:
   - Verificará Python
   - Creará entorno virtual
   - Instalará dependencias
   - Abrirá el navegador

### Método 2: Manual

```bash
# 1. Crear entorno virtual
python -m venv venv

# 2. Activar entorno virtual (Windows)
venv\Scripts\activate

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Configurar credenciales (ver sección siguiente)

# 5. Iniciar servidor
python app.py
```

## Configuración de Correo SMTP

### Para Gmail

1. Ve a tu cuenta de Google: [myaccount.google.com/security](https://myaccount.google.com/security)
2. Activa **Verificación en 2 pasos**
3. Busca **Contraseñas de aplicaciones**
4. Genera una contraseña para "Correo" y "Windows"
5. Copia el archivo `.env.example` a `.env`:
   ```bash
   copy .env.example .env
   ```
6. Edita `.env` y configura:
   ```env
   EMAIL_HOST=smtp.gmail.com
   EMAIL_PORT=587
   EMAIL_USER=tu_correo@gmail.com
   EMAIL_PASSWORD=xxxx xxxx xxxx xxxx  # Contraseña de aplicación de 16 dígitos
   EMAIL_FROM_NAME=Cartera Lomarosa
   ```

### Para Outlook/Office 365

1. Copia `.env.example` a `.env`
2. Edita `.env`:
   ```env
   EMAIL_HOST=smtp-mail.outlook.com  # O smtp.office365.com para corporativo
   EMAIL_PORT=587
   EMAIL_USER=tu_correo@outlook.com
   EMAIL_PASSWORD=tu_contraseña
   EMAIL_FROM_NAME=Cartera Lomarosa
   ```

## Cómo Usar

### Paso 1: Ejecutar la aplicación

Haz doble clic en **`iniciar.bat`** o ejecuta:
```bash
python app.py
```

El navegador se abrirá automáticamente en `http://localhost:5000`

### Paso 2: Cargar archivos Excel

Arrastra y suelta (o haz clic para seleccionar) dos archivos:

1. **Listado de Terceros**
   - Debe contener columnas: `NIT`, `Correo`, `Condición de pago`
   - La condición de pago puede ser: "Contado", "15 dias", "30 dias", "Treinta dias", etc.

2. **Cartera por Edades**
   - Debe contener: `NIT`, `Factura`, `Fecha Vencimiento`, `Días`, `Saldo`

### Paso 3: Analizar archivos

Haz clic en **"Analizar Archivos"**

El sistema:
- Detecta automáticamente qué fila es el encabezado real
- Procesa las facturas que están próximas a vencer (< 5 días)
- Procesa las facturas que ya vencieron
- Muestra estadísticas y tabla con vista previa

### Paso 4: Revisar vista previa

Verás:
- **Próximos a vencer**: Facturas con menos de 5 días para vencer
- **Vencidos**: Facturas que ya pasaron su fecha de vencimiento
- **Total a enviar**: Cantidad de correos que se enviarán
- Tabla completa con detalles de cada recordatorio

### Paso 5: Enviar correos

Haz clic en **"Enviar Correos Ahora"**

El sistema:
- Enviará correos en paralelo (5 simultáneos por defecto)
- Mostrará barra de progreso
- Presentará resultados detallados (exitosos y fallidos)

## Estructura del Proyecto

```
cartera_final/
│
├── app.py                     # Backend Flask con envío SMTP
├── iniciar.bat                # Script de inicio automático (Windows)
├── requirements.txt           # Dependencias de Python
├── .env.example              # Plantilla de configuración
├── .env                      # Configuración real (crear manualmente)
├── README.md                 # Esta documentación
│
├── templates/
│   └── index.html            # Interfaz web (3 pasos)
│
└── static/
    ├── css/
    │   └── styles.css        # Estilos modernos con gradientes
    └── js/
        └── app.js            # Lógica drag & drop + procesamiento Excel
```

## Archivos Excel: Columnas Esperadas

### Listado de Terceros

El sistema detecta automáticamente el encabezado. Debe tener:

- **NIT / Cédula**: Identificación del tercero
- **Condición de pago**: "Contado", "15 dias", "30 dias", "Treinta dias", etc.
- **Correo / Mail contacto**: Email del destinatario

### Cartera por Edades

El sistema detecta automáticamente el encabezado. Debe tener:

- **NIT / Tercero**: Identificación del tercero
- **Factura / Documento / Numero FAC**: Número de factura
- **Vencimiento / Fecha**: Fecha de vencimiento
- **Dias**: Días de vencimiento (opcional)
- **Saldo / Valor**: Monto de la factura

## Lógica de Recordatorios

El sistema incluye facturas que:

1. **Ya vencieron** (días desde vencimiento < 0)
2. **Están próximas a vencer** (menos de 5 días)

**NO incluye**:
- Facturas con más de 5 días de anticipación
- Terceros sin correo electrónico
- Facturas sin fecha de vencimiento válida

## Probar Configuración SMTP

Para verificar que tu configuración de correo funciona:

1. Abre en el navegador: `http://localhost:5000/test-email`
2. Deberías recibir un correo de prueba en tu bandeja de entrada

## Solución de Problemas

### El archivo iniciar.bat no ejecuta

**Solución:**
- Verifica que Python esté instalado: abre CMD y escribe `python --version`
- Instala Python desde [python.org](https://www.python.org/downloads/)
- Marca la opción "Add Python to PATH" durante la instalación

### Los correos no se envían

**Soluciones:**
1. Verifica que el archivo `.env` existe y tiene las credenciales correctas
2. Para Gmail: usa una **contraseña de aplicación** (NO tu contraseña normal)
3. Prueba primero con `/test-email` en el navegador
4. Revisa que no haya espacios extra en el correo o contraseña

### Error al leer Excel

**Soluciones:**
- Verifica que el archivo sea `.xlsx` o `.xls`
- Asegúrate de que las columnas existan (el sistema detecta automáticamente)
- Revisa la consola del navegador (F12) para ver el error específico

### No se encuentran las columnas en Excel

**Soluciones:**
- El sistema busca automáticamente las columnas por nombre (case-insensitive)
- Acepta variaciones: "NIT", "Cédula", "cedula", "Mail", "Correo", etc.
- Si el Excel tiene filas de instrucciones antes del encabezado, el sistema las ignora automáticamente

### Firewall bloquea el puerto 587

**Solución:**
- Configura tu firewall para permitir conexiones salientes al puerto 587
- En redes corporativas, contacta al administrador de TI

## Clonar en Otra PC

1. Copia la carpeta completa `cartera_final` a la otra PC
2. Asegúrate de que Python esté instalado
3. Ejecuta `iniciar.bat`
4. Configura el archivo `.env` con las credenciales de correo

## Variables de Entorno (.env)

| Variable | Descripción | Valor por defecto |
|----------|-------------|-------------------|
| `EMAIL_HOST` | Servidor SMTP | `smtp.gmail.com` |
| `EMAIL_PORT` | Puerto SMTP | `587` |
| `EMAIL_USER` | Correo remitente | (requerido) |
| `EMAIL_PASSWORD` | Contraseña de aplicación | (requerido) |
| `EMAIL_FROM_NAME` | Nombre que aparece en "De:" | `Cartera Lomarosa` |
| `MAX_WORKERS` | Correos enviados en paralelo | `5` |

## Características Técnicas

- **Backend**: Flask + SMTP (smtplib)
- **Frontend**: HTML5 + CSS3 + Vanilla JavaScript
- **Librerías JS**: XLSX.js (lectura Excel) + Day.js (fechas)
- **Envío paralelo**: ThreadPoolExecutor (Python)
- **Sin dependencias externas** para procesamiento de Excel (todo en el navegador)

## Licencia

Proyecto interno - Cartera Lomarosa

## Soporte

Para problemas o preguntas, contacta al departamento de TI o Cartera.
