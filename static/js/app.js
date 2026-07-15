// ==========================================
// VARIABLES GLOBALES
// ==========================================

let file1Obj = null;
let file2Obj = null;
let recordatoriosGlobal = [];
let clientesAgrupados = [];
let fechaCarteraGlobal = null;
let infoCierreGlobal = null;
let enviados_log = [];
let ciclo_log = null;

// ==========================================
// UTILIDADES
// ==========================================

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

function normalizarTexto(texto) {
    if (!texto) return "";
    return String(texto).trim().toLowerCase();
}

// ==========================================
// DRAG & DROP
// ==========================================

function setupDragAndDrop() {
    const dropZoneTerceros = document.getElementById("dropZoneTerceros");
    const dropZoneCartera = document.getElementById("dropZoneCartera");
    const fileTercerosInput = document.getElementById("fileTerceros");
    const fileCarteraInput = document.getElementById("fileCartera");

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ["dragenter", "dragover", "dragleave", "drop"].forEach(eventName => {
        dropZoneTerceros.addEventListener(eventName, preventDefaults, false);
    });

    ["dragenter", "dragover"].forEach(eventName => {
        dropZoneTerceros.addEventListener(eventName, () => {
            dropZoneTerceros.classList.add("dragover");
        });
    });

    ["dragleave", "drop"].forEach(eventName => {
        dropZoneTerceros.addEventListener(eventName, () => {
            dropZoneTerceros.classList.remove("dragover");
        });
    });

    dropZoneTerceros.addEventListener("drop", e => {
        const files = e.dataTransfer.files;
        if (files.length > 0) handleFile1(files[0]);
    });

    fileTercerosInput.addEventListener("change", e => {
        if (e.target.files.length > 0) handleFile1(e.target.files[0]);
    });

    ["dragenter", "dragover", "dragleave", "drop"].forEach(eventName => {
        dropZoneCartera.addEventListener(eventName, preventDefaults, false);
    });

    ["dragenter", "dragover"].forEach(eventName => {
        dropZoneCartera.addEventListener(eventName, () => {
            dropZoneCartera.classList.add("dragover");
        });
    });

    ["dragleave", "drop"].forEach(eventName => {
        dropZoneCartera.addEventListener(eventName, () => {
            dropZoneCartera.classList.remove("dragover");
        });
    });

    dropZoneCartera.addEventListener("drop", e => {
        const files = e.dataTransfer.files;
        if (files.length > 0) handleFile2(files[0]);
    });

    fileCarteraInput.addEventListener("change", e => {
        if (e.target.files.length > 0) handleFile2(e.target.files[0]);
    });
}

function handleFile1(file) {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        alert("Por favor selecciona un archivo Excel (.xlsx o .xls)");
        return;
    }
    file1Obj = file;
    const info = document.getElementById("infoTerceros");
    info.innerHTML = `<strong>✓ Archivo cargado:</strong><br>${file.name}<br><small>${formatFileSize(file.size)}</small>`;
    info.style.display = "block";
    document.getElementById("dropZoneTerceros").classList.add("file-loaded");
    checkFilesReady();
}

function handleFile2(file) {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        alert("Por favor selecciona un archivo Excel (.xlsx o .xls)");
        return;
    }
    file2Obj = file;
    const info = document.getElementById("infoCartera");
    info.innerHTML = `<strong>✓ Archivo cargado:</strong><br>${file.name}<br><small>${formatFileSize(file.size)}</small>`;
    info.style.display = "block";
    document.getElementById("dropZoneCartera").classList.add("file-loaded");
    checkFilesReady();
}

function checkFilesReady() {
    const btnAnalizar = document.getElementById("btnAnalizar");
    btnAnalizar.disabled = !(file1Obj && file2Obj);
}

// ==========================================
// AGRUPAR POR CLIENTE (UNIFICADO)
// ==========================================

function agruparPorCliente(recordatorios) {
    const agrupados = {};

    recordatorios.forEach(r => {
        const email = r.correo_cliente;
        const cliente = r.cliente;
        const estado = r.estado;

        // KEY única: cliente + email
        const key = `${cliente}|${email}`;

        if (!agrupados[key]) {
            agrupados[key] = {
                cliente: cliente,
                correo_cliente: email,
                vendedor: r.vendedor,
                correo_vendedor: r.correo_vendedor,
                local: r.local,
                facturas_vencidas: [],
                facturas_proximas: [],
                facturas_no_vencidas: [],
                total_facturas: 0,
                total_vencidas: 0,
                total_proximas: 0,
                total_no_vencidas: 0,
                total_saldo: 0,
                cupo: r.cupo || 0,
                cupo_disponible: 0
            };
        }

        // Clasificar factura según estado
        const factura = {
            numero_factura: r.numero_factura,
            fecha_emision: r.fecha_emision,
            fecha_vencimiento: r.fecha_vencimiento,
            dias: r.dias,
            saldo: r.saldo,
            saldo_numerico: r.saldo_numerico,
            estado: r.estado,
            correo_cliente: r.correo_cliente,
            correo_vendedor: r.correo_vendedor,
            local: r.local
        };

        if (estado === 'vencido') {
            agrupados[key].facturas_vencidas.push(factura);
            agrupados[key].total_vencidas += 1;
        } else if (estado === 'proximo') {
            agrupados[key].facturas_proximas.push(factura);
            agrupados[key].total_proximas += 1;
        } else if (estado === 'no_vencido') {
            agrupados[key].facturas_no_vencidas.push(factura);
            agrupados[key].total_no_vencidas += 1;
        }

        agrupados[key].total_facturas += 1;
        agrupados[key].total_saldo += r.saldo_numerico || 0;
    });

    // Calcular cupo_disponible para cada cliente
    Object.values(agrupados).forEach(cliente => {
        cliente.cupo_disponible = cliente.cupo - cliente.total_saldo;
    });

    return Object.values(agrupados).sort((a, b) => {
        // Ordenar por total de vencidas (descendente)
        return b.total_vencidas - a.total_vencidas;
    });
}

// ==========================================
// SISTEMA DE TANDAS (LOG DE ENVIADOS)
// ==========================================

async function cargarLogEnviados() {
  try {
    const resp = await fetch('/log-enviados');
    const data = await resp.json();
    enviados_log = data.enviados || [];
    ciclo_log = data.ciclo;
    actualizarBannerCiclo();
  } catch(e) {
    enviados_log = [];
  }
}

function actualizarBannerCiclo() {
  const banner = document.getElementById('banner-ciclo');
  if (!banner) return;

  const clientesUnicos = clientesAgrupados ? clientesAgrupados.length : 0;
  const yaEnviados = clientesAgrupados ? clientesAgrupados.filter(c => enviados_log.includes(c.correo_cliente)).length : 0;
  const pendientes = clientesUnicos - yaEnviados;

  document.getElementById('ciclo-actual').textContent = ciclo_log || '';
  document.getElementById('ya-enviados').textContent = yaEnviados;
  document.getElementById('pendientes-count').textContent = pendientes;
  banner.style.display = yaEnviados > 0 ? 'flex' : 'none';
}

function seleccionarPendientes() {
  const limite = parseInt(document.getElementById('limite-tanda').value) || 90;
  const checkboxes = document.querySelectorAll('.check-cliente');
  let seleccionados = 0;

  // Primero desmarcar todos
  checkboxes.forEach(cb => {
    cb.checked = false;
    cb.disabled = false;
  });

  // Marcar solo pendientes hasta el límite
  checkboxes.forEach(cb => {
    const email = cb.dataset.email;
    if (!enviados_log.includes(email) && seleccionados < limite) {
      cb.checked = true;
      seleccionados++;
    }
  });

  actualizarConteoEnvio();
}

async function resetearCiclo() {
  if (!confirm('¿Iniciar un nuevo ciclo? Esto borrará el historial de enviados del ciclo actual.')) return;
  await fetch('/reset-log', { method: 'POST' });
  await cargarLogEnviados();
  renderTablaUnificada();
}

// ==========================================
// ANALIZAR ARCHIVOS
// ==========================================

async function analizarArchivos() {
    const btnAnalizar = document.getElementById("btnAnalizar");
    btnAnalizar.disabled = true;
    btnAnalizar.textContent = "Procesando...";

    try {
        const formData = new FormData();
        formData.append("file1", file1Obj);
        formData.append("file2", file2Obj);

        const response = await fetch("/procesar-excel", {
            method: "POST",
            body: formData
        });

        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.message || "Error al procesar archivos");
        }

        const resultado = await response.json();

        if (!resultado.success) {
            throw new Error(resultado.message || "Error desconocido");
        }

        recordatoriosGlobal = resultado.recordatorios || [];
        fechaCarteraGlobal = resultado.fecha_cartera || null;
        infoCierreGlobal = resultado.info_cierre || null;

        console.log(`📅 Fecha Cartera: ${fechaCarteraGlobal}`);
        console.log(`📋 Info Cierre:`, infoCierreGlobal);

        if (recordatoriosGlobal.length === 0) {
            alert("No se encontraron facturas con email asignado.");
            btnAnalizar.disabled = false;
            btnAnalizar.textContent = "Analizar Archivos";
            return;
        }

        // Agrupar clientes de forma unificada
        clientesAgrupados = agruparPorCliente(recordatoriosGlobal);

        renderTablaUnificada();
        renderEstadisticas(resultado.stats);

        // Mostrar fecha cartera
        mostrarFechaCartera(fechaCarteraGlobal);

        // Mostrar alerta de cierre trimestral si aplica
        mostrarAlertaCierre(infoCierreGlobal, fechaCarteraGlobal);

        document.getElementById("step2").style.display = "block";
        document.getElementById("step3").style.display = "block";

        document.getElementById("step2").scrollIntoView({ behavior: "smooth" });

        // Cargar log de enviados y actualizar banner de ciclo
        await cargarLogEnviados();

        btnAnalizar.textContent = "Analizar Archivos";
        btnAnalizar.disabled = false;

    } catch (error) {
        console.error("Error:", error);
        alert("Error al procesar los archivos:\n\n" + error.message);
        btnAnalizar.disabled = false;
        btnAnalizar.textContent = "Analizar Archivos";
    }
}

// ==========================================
// RENDER DE TABLA UNIFICADA
// ==========================================

function renderTablaUnificada() {
    const filterValue = document.getElementById("filterClientes").value.toLowerCase();
    const tbody = document.getElementById("tbodyClientes");
    tbody.innerHTML = '';

    const filtrados = clientesAgrupados.filter(c =>
        c.cliente.toLowerCase().includes(filterValue)
    );

    document.getElementById("countClientesTabla").textContent = filtrados.length;

    filtrados.forEach((cliente, idx) => {
        const uniqueKey = `${cliente.cliente}|${cliente.correo_cliente}`;
        const yaEnviado = enviados_log.includes(cliente.correo_cliente);

        // Formatear montos
        const totalCarteraFormat = `$${cliente.total_saldo.toLocaleString('es-CO', {maximumFractionDigits: 0})}`;
        const cupoDisponibleFormat = `$${cliente.cupo_disponible.toLocaleString('es-CO', {maximumFractionDigits: 0})}`;
        const cupoDisponibleColor = cliente.cupo_disponible < 0 ? '#dc2626' : '#10b981';

        const badgeEnviado = yaEnviado
            ? `<span style="background:#dcfce7; color:#16a34a; font-size:11px; font-weight:bold; padding:2px 7px; border-radius:999px; margin-left:6px; vertical-align:middle;">Enviado</span>`
            : '';

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>
                <input type="checkbox" class="check-cliente cliente-checkbox" value="${uniqueKey}"
                       data-cliente="${cliente.cliente}"
                       data-email="${cliente.correo_cliente}"
                       ${yaEnviado ? '' : 'checked'}
                       ${yaEnviado ? 'disabled' : ''}>
            </td>
            <td><strong>${cliente.cliente}</strong>${badgeEnviado}</td>
            <td>${cliente.correo_cliente}</td>
            <td>${cliente.vendedor}</td>
            <td style="text-align: center; font-weight: bold;">${cliente.total_facturas}</td>
            <td style="text-align: center; color: #dc2626; font-weight: bold;">${cliente.total_vencidas}</td>
            <td style="text-align: center; color: #f59e0b; font-weight: bold;">${cliente.total_proximas}</td>
            <td style="text-align: center; color: #10b981; font-weight: bold;">${cliente.total_no_vencidas}</td>
            <td style="text-align: right; font-weight: bold; color: #dc2626;">${totalCarteraFormat}</td>
            <td style="text-align: right; font-weight: bold; color: ${cupoDisponibleColor};">${cupoDisponibleFormat}</td>
            <td>
                <button class="btn-expand" onclick="toggleFacturasCliente(${idx})">
                    <span id="expand-cliente-${idx}">▼</span> Ver Facturas
                </button>
            </td>
        `;

        tbody.appendChild(tr);

        // Fila de detalle con TRES sub-tablas
        const detailRow = document.createElement('tr');
        detailRow.id = `detail-cliente-${idx}`;
        detailRow.style.display = 'none';

        // Generar sub-tabla de vencidas
        const subTablaVencidas = generarSubTabla(
            cliente.facturas_vencidas,
            "VENCIDAS",
            "#dc2626",
            "🔴"
        );

        // Generar sub-tabla de próximas
        const subTablaProximas = generarSubTabla(
            cliente.facturas_proximas,
            "PRÓXIMAS (≤ 5 días)",
            "#f59e0b",
            "🟡"
        );

        // Generar sub-tabla de no vencidas
        const subTablaNoVencidas = generarSubTabla(
            cliente.facturas_no_vencidas,
            "NO VENCIDAS (> 5 días)",
            "#10b981",
            "🟢"
        );

        detailRow.innerHTML = `
            <td colspan="11" style="padding: 20px; background-color: #f8f9fa;">
                <div style="display: grid; gap: 20px;">
                    ${subTablaVencidas}
                    ${subTablaProximas}
                    ${subTablaNoVencidas}
                </div>
            </td>
        `;

        tbody.appendChild(detailRow);
    });

    actualizarConteoEnvio();
}

function generarSubTabla(facturas, titulo, colorBg, emoji) {
    if (!facturas || facturas.length === 0) {
        return '';
    }

    const filas = facturas.map(f => `
        <div style="display: grid; grid-template-columns: 120px 100px 100px 80px 150px 150px 100px; gap: 10px; padding: 10px 0; border-bottom: 1px solid #eee; font-size: 13px;">
            <div><strong>${f.numero_factura}</strong></div>
            <div>${f.fecha_emision}</div>
            <div>${f.fecha_vencimiento}</div>
            <div>${f.dias} días</div>
            <div style="text-align: right; font-weight: bold;">${f.saldo}</div>
            <div>${f.correo_cliente}</div>
            <div>${f.local || 'N/A'}</div>
        </div>
    `).join('');

    return `
        <div style="border: 2px solid ${colorBg}; border-radius: 8px; padding: 15px; background: white;">
            <h4 style="margin: 0 0 15px 0; color: ${colorBg}; border-bottom: 2px solid ${colorBg}; padding-bottom: 8px;">
                ${emoji} ${titulo} (${facturas.length})
            </h4>
            <div style="display: grid; grid-template-columns: 120px 100px 100px 80px 150px 150px 100px; gap: 10px; margin-bottom: 10px; font-weight: bold; color: #666; font-size: 12px;">
                <div>Factura</div>
                <div>Emisión</div>
                <div>Vencimiento</div>
                <div>Días</div>
                <div style="text-align: right;">Saldo</div>
                <div>Email</div>
                <div>Local</div>
            </div>
            ${filas}
        </div>
    `;
}

function toggleFacturasCliente(idx) {
    const detailRow = document.getElementById(`detail-cliente-${idx}`);
    const expandIcon = document.getElementById(`expand-cliente-${idx}`);

    if (detailRow.style.display === 'none') {
        detailRow.style.display = 'table-row';
        expandIcon.textContent = '▲';
    } else {
        detailRow.style.display = 'none';
        expandIcon.textContent = '▼';
    }
}

function renderEstadisticas(stats) {
    document.getElementById("statVencidas").textContent = stats.vencidas || 0;
    document.getElementById("statProximas").textContent = stats.proximas || 0;
    document.getElementById("statNoVencidas").textContent = stats.no_vencidas || 0;
    document.getElementById("statTotal").textContent = stats.total || 0;
}

// ==========================================
// FUNCIONES DE CIERRE TRIMESTRAL
// ==========================================

function mostrarFechaCartera(fechaCartera) {
    const infoDiv = document.getElementById("fechaCarteraInfo");
    const fechaDisplay = document.getElementById("fechaCarteraDisplay");

    if (fechaCartera) {
        fechaDisplay.textContent = fechaCartera;
        infoDiv.style.display = "block";
    } else {
        infoDiv.style.display = "none";
    }
}

function mostrarAlertaCierre(infoCierre, fechaCartera) {
    const alertDiv = document.getElementById("cierreAlert");
    const tituloEl = document.getElementById("cierreTitulo");
    const fechaEl = document.getElementById("cierreFecha");
    const checkbox = document.getElementById("checkIncluirMensajeCierre");

    if (infoCierre && infoCierre.es_cierre) {
        // Determinar título según tipo de cierre
        const tipoTexto = infoCierre.tipo === "anual" ? "🎄 Cierre Anual Detectado" : "📋 Cierre Trimestral Detectado";
        tituloEl.textContent = tipoTexto;
        fechaEl.textContent = infoCierre.fecha_formateada || fechaCartera;

        // Activar checkbox por defecto
        checkbox.checked = true;

        // Mostrar alerta
        alertDiv.style.display = "block";

        console.log(`✅ Cierre ${infoCierre.tipo} detectado para fecha ${infoCierre.fecha_formateada}`);
    } else {
        alertDiv.style.display = "none";
    }
}

function actualizarConteoEnvio() {
    const clientesSeleccionados = document.querySelectorAll('.check-cliente:checked').length;
    document.getElementById("countClientesEnviar").textContent = clientesSeleccionados;
}

// ==========================================
// FILTRO DE BÚSQUEDA
// ==========================================

document.addEventListener("DOMContentLoaded", () => {
    setupDragAndDrop();

    // Filtro
    const filterInput = document.getElementById("filterClientes");
    if (filterInput) {
        filterInput.addEventListener("input", () => {
            renderTablaUnificada();
        });
    }

    // Checkbox "Seleccionar todos"
    const selectAllClientes = document.getElementById("selectAllClientes");
    if (selectAllClientes) {
        selectAllClientes.addEventListener("change", (e) => {
            document.querySelectorAll('.check-cliente').forEach(cb => cb.checked = e.target.checked);
            actualizarConteoEnvio();
        });
    }

    // Cambio en checkboxes individuales
    document.addEventListener("change", (e) => {
        if (e.target.classList.contains('check-cliente')) {
            actualizarConteoEnvio();
        }
    });

    document.getElementById("btnAnalizar").addEventListener("click", analizarArchivos);
    document.getElementById("btnEnviarCorreos").addEventListener("click", enviarCorreos);

    console.log("✅ App inicializada");
});

// ==========================================
// ENVÍO DE CORREOS UNIFICADO
// ==========================================

async function enviarCorreos() {
    const checkboxes = document.querySelectorAll('.check-cliente:checked');

    if (checkboxes.length === 0) {
        alert("No hay clientes seleccionados");
        return;
    }

    // Verificar si se incluye mensaje de cierre
    const checkIncluirCierre = document.getElementById("checkIncluirMensajeCierre");
    const incluirMensajeCierre = checkIncluirCierre ? checkIncluirCierre.checked : false;

    // Construir mensaje de confirmación
    let mensajeConfirm = `¿Enviar ${checkboxes.length} correos unificados?\n\n`;
    mensajeConfirm += `Cada cliente recibirá UN SOLO correo con todas sus facturas (vencidas, próximas y no vencidas).`;

    if (infoCierreGlobal && infoCierreGlobal.es_cierre) {
        if (incluirMensajeCierre) {
            const tipoCierre = infoCierreGlobal.tipo === "anual" ? "CIERRE ANUAL" : "CIERRE TRIMESTRAL";
            mensajeConfirm += `\n\n📋 Se incluirá mensaje de ${tipoCierre} en los correos.`;
        } else {
            mensajeConfirm += `\n\n⚠️ El mensaje de cierre trimestral NO se incluirá.`;
        }
    }

    const confirmacion = confirm(mensajeConfirm);

    if (!confirmacion) {
        return;
    }

    // Extraer clientes seleccionados
    const clientesSeleccionados = Array.from(checkboxes).map(cb => ({
        cliente: cb.getAttribute('data-cliente'),
        email: cb.getAttribute('data-email')
    }));

    console.log(`📧 Enviando correos a ${clientesSeleccionados.length} clientes...`);
    console.log(`📋 Incluir mensaje cierre: ${incluirMensajeCierre}`);

    // Filtrar recordatorios para clientes seleccionados
    const recordatoriosFiltrados = recordatoriosGlobal.filter(r => {
        return clientesSeleccionados.some(cs => {
            const clienteMatch = normalizarTexto(cs.cliente) === normalizarTexto(r.cliente);
            const emailMatch = normalizarTexto(cs.email) === normalizarTexto(r.correo_cliente);
            return clienteMatch && emailMatch;
        });
    });

    console.log(`Total recordatorios a enviar: ${recordatoriosFiltrados.length}`);

    const btn = document.getElementById("btnEnviarCorreos");
    btn.disabled = true;
    btn.textContent = "Enviando...";

    const progressArea = document.getElementById("progressArea");
    progressArea.style.display = "block";

    const progressFill = document.getElementById("progressFill");
    const progressText = document.getElementById("progressText");

    try {
        progressFill.style.width = "20%";
        progressText.textContent = `Enviando ${clientesSeleccionados.length} correos...`;

        const response = await fetch("/enviar-correos", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                recordatorios: recordatoriosFiltrados,
                fecha_cartera: fechaCarteraGlobal,
                info_cierre: infoCierreGlobal,
                incluir_mensaje_cierre: incluirMensajeCierre
            })
        });

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.message || "Error en servidor");
        }

        const resultado = await response.json();

        console.log(`✅ Envío completado:`);
        console.log(`  - Total: ${resultado.total}`);
        console.log(`  - Exitosos: ${resultado.exitosos}`);
        console.log(`  - Fallidos: ${resultado.fallidos}`);

        progressFill.style.width = "100%";
        progressText.textContent = "✅ Envío completado";

        document.getElementById("resultExitosos").textContent = resultado.exitosos;
        document.getElementById("resultFallidos").textContent = resultado.fallidos;

        document.getElementById("resultsArea").style.display = "block";

        setTimeout(() => progressArea.style.display = "none", 2000);

    } catch (error) {
        console.error(`❌ ERROR EN ENVÍO:`, error);
        alert("Error: " + error.message);
        progressText.textContent = "❌ Error en envío";
    } finally {
        btn.disabled = false;
        btn.textContent = "Enviar Correos";
    }
}
