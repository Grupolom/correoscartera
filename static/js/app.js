// ==========================================
// VARIABLES GLOBALES
// ==========================================

let file1Obj = null;
let file2Obj = null;
let recordatoriosGlobal = [];
let clientesVencidos = [];
let clientesProximos = [];

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
    /**
     * Normaliza texto para comparaciones robustas
     * - Convierte a lowercase
     * - Elimina espacios al inicio y final
     * - Maneja valores null/undefined
     */
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
// AGRUPAR Y SEPARAR POR ESTADO
// ==========================================

function agruparPorClienteYEstado(recordatorios) {
    const vencidos = {};
    const proximos = {};

    recordatorios.forEach(r => {
        const email = r.correo_cliente;
        const cliente = r.cliente;
        const estado = r.estado;
        const target = estado === 'vencido' ? vencidos : proximos;

        // KEY única: cliente + email (permite separar clientes con mismo email)
        const key = `${cliente}|${email}`;

        if (!target[key]) {
            target[key] = {
                cliente: cliente,
                correo_cliente: email,
                vendedor: r.vendedor,
                correo_vendedor: r.correo_vendedor,
                local: r.local,
                facturas: []
            };
        }

        target[key].facturas.push({
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
        });
    });

    // ← DEBUG: Verificar separación de clientes con mismo email
    const abelVencidos = Object.values(vencidos).filter(c => c.cliente.toLowerCase().includes('abel'));
    if (abelVencidos.length > 0) {
        console.log(`🔍 DEBUG - Clientes con "abel" en nombre (VENCIDOS):`);
        abelVencidos.forEach(c => {
            console.log(`  Cliente: ${c.cliente}`);
            console.log(`  Email: ${c.correo_cliente}`);
            console.log(`  Facturas: ${c.facturas.length}`);
            console.log(`  Números:`, c.facturas.map(f => f.numero_factura));
        });
    }

    return {
        vencidos: Object.values(vencidos).sort((a, b) => {
            const dateA = new Date(a.facturas[0].fecha_vencimiento.split('/').reverse().join('-'));
            const dateB = new Date(b.facturas[0].fecha_vencimiento.split('/').reverse().join('-'));
            return dateA - dateB;
        }),
        proximos: Object.values(proximos).sort((a, b) => {
            const dateA = new Date(a.facturas[0].fecha_vencimiento.split('/').reverse().join('-'));
            const dateB = new Date(b.facturas[0].fecha_vencimiento.split('/').reverse().join('-'));
            return dateA - dateB;
        })
    };
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

        if (recordatoriosGlobal.length === 0) {
            alert("No se encontraron facturas próximas a vencer o vencidas con email asignado.");
            btnAnalizar.disabled = false;
            btnAnalizar.textContent = "Analizar Archivos";
            return;
        }

        const agrupado = agruparPorClienteYEstado(recordatoriosGlobal);
        clientesVencidos = agrupado.vencidos;
        clientesProximos = agrupado.proximos;

        renderTablas();
        renderEstadisticas(resultado.stats);

        document.getElementById("step2").style.display = "block";
        document.getElementById("step3").style.display = "block";
        document.getElementById("step4").style.display = "block";

        document.getElementById("step2").scrollIntoView({ behavior: "smooth" });

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
// RENDER DE TABLAS CON EXPANDIBLES
// ==========================================

function renderTablas() {
    const filterValue = document.getElementById("filterClientes").value.toLowerCase();
    
    // Tablas Vencidos
    renderTablaEstado('vencidos', clientesVencidos, filterValue);
    
    // Tabla Próximos
    renderTablaEstado('proximos', clientesProximos, filterValue);

    // Actualizar conteos
    document.getElementById("countVencidosTabla").textContent = clientesVencidos.filter(c => 
        c.cliente.toLowerCase().includes(filterValue)
    ).length;
    
    document.getElementById("countProximosTabla").textContent = clientesProximos.filter(c => 
        c.cliente.toLowerCase().includes(filterValue)
    ).length;

    actualizarConteoEnvio();
}

function renderTablaEstado(estado, clientes, filterValue = '') {
    const tableId = estado === 'vencidos' ? 'tablaVencidos' : 'tablaProximos';
    const tbody = document.getElementById(`tbody${estado.charAt(0).toUpperCase() + estado.slice(1)}`);
    tbody.innerHTML = '';

    const filtrados = clientes.filter(c => c.cliente.toLowerCase().includes(filterValue));

    filtrados.forEach((cliente, idx) => {
        // KEY única para identificar cada cliente (incluye nombre para separar clientes con mismo email)
        const uniqueKey = `${cliente.cliente}|${cliente.correo_cliente}`;

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>
                <input type="checkbox" class="check-${estado}" value="${uniqueKey}" data-cliente="${cliente.cliente}" data-email="${cliente.correo_cliente}" checked>
            </td>
            <td><strong>${cliente.cliente}</strong></td>
            <td>${cliente.correo_cliente}</td>
            <td>${cliente.vendedor}</td>
            <td style="text-align: center; font-weight: bold;">${cliente.facturas.length}</td>
            <td>
                <button class="btn-expand" onclick="toggleFacturas('${estado}', ${idx})">
                    <span id="expand-${estado}-${idx}">▼</span> Ver Facturas
                </button>
            </td>
        `;

        tbody.appendChild(tr);

        // Fila de detalle CON MÁS DATOS
        const detailRow = document.createElement('tr');
        detailRow.id = `detail-${estado}-${idx}`;
        detailRow.style.display = 'none';
        
        const facturasList = cliente.facturas.map(f => `
            <div style="display: grid; grid-template-columns: 100px 100px 100px 120px 80px 180px 180px 100px; gap: 10px; padding: 10px 0; border-bottom: 1px solid #eee; font-size: 13px;">
                <div><strong>${f.numero_factura}</strong></div>
                <div>${f.fecha_emision}</div>
                <div>${f.fecha_vencimiento}</div>
                <div style="text-align: right;">${f.saldo}</div>
                <div>${f.dias} días</div>
                <div>${f.correo_cliente || cliente.correo_cliente}</div>
                <div>${f.correo_vendedor || cliente.correo_vendedor}</div>
                <div>${f.local || cliente.local}</div>
            </div>
        `).join('');

        detailRow.innerHTML = `
            <td colspan="6" style="padding: 20px;">
                <div style="background: #f8f9fa; padding: 15px; border-radius: 8px;">
                    <div style="display: grid; grid-template-columns: 100px 100px 100px 120px 80px 180px 180px 100px; gap: 10px; margin-bottom: 10px; font-weight: bold; color: #666; font-size: 12px;">
                        <div>Factura</div>
                        <div>Emisión</div>
                        <div>Vencimiento</div>
                        <div style="text-align: right;">Saldo</div>
                        <div>Días</div>
                        <div>Email Cliente</div>
                        <div>Email Vendedor</div>
                        <div>Local</div>
                    </div>
                    ${facturasList}
                </div>
            </td>
        `;

        tbody.appendChild(detailRow);
    });
}


function toggleFacturas(estado, idx) {
    const detailRow = document.getElementById(`detail-${estado}-${idx}`);
    const expandIcon = document.getElementById(`expand-${estado}-${idx}`);
    
    if (detailRow.style.display === 'none') {
        detailRow.style.display = 'table-row';
        expandIcon.textContent = '▲';
    } else {
        detailRow.style.display = 'none';
        expandIcon.textContent = '▼';
    }
}

function renderEstadisticas(stats) {
    document.getElementById("statProximos").textContent = stats.proximos || 0;
    document.getElementById("statVencidos").textContent = stats.vencidos || 0;
    document.getElementById("statTotal").textContent = (stats.vencidos + stats.proximos) || 0;
}

function actualizarConteoEnvio() {
    const vencidosSeleccionados = document.querySelectorAll('.check-vencidos:checked').length;
    const proximosSeleccionados = document.querySelectorAll('.check-proximos:checked').length;
    
    document.getElementById("countVencidosEnviar").textContent = vencidosSeleccionados;
    document.getElementById("countProximosEnviar").textContent = proximosSeleccionados;
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
            renderTablas();
        });
    }

    // Checkboxes "Seleccionar todos"
    const selectAllVencidos = document.getElementById("selectAllVencidos");
    const selectAllProximos = document.getElementById("selectAllProximos");

    if (selectAllVencidos) {
        selectAllVencidos.addEventListener("change", (e) => {
            document.querySelectorAll('.check-vencidos').forEach(cb => cb.checked = e.target.checked);
            actualizarConteoEnvio();
        });
    }

    if (selectAllProximos) {
        selectAllProximos.addEventListener("change", (e) => {
            document.querySelectorAll('.check-proximos').forEach(cb => cb.checked = e.target.checked);
            actualizarConteoEnvio();
        });
    }

    // Cambio en checkboxes individuales
    document.addEventListener("change", (e) => {
        if (e.target.classList.contains('check-vencidos') || e.target.classList.contains('check-proximos')) {
            actualizarConteoEnvio();
        }
    });

    document.getElementById("btnAnalizar").addEventListener("click", analizarArchivos);
    document.getElementById("btnEnviarVencidos").addEventListener("click", () => enviarCorreosEstado('vencidos'));
    document.getElementById("btnEnviarProximos").addEventListener("click", () => enviarCorreosEstado('proximos'));

    console.log("✅ App inicializada");
});

// ==========================================
// ENVÍO DE CORREOS POR ESTADO
// ==========================================

async function enviarCorreosEstado(estado) {
    const checkboxes = document.querySelectorAll(`.check-${estado}:checked`);

    // Normalizar estado: "vencidos" → "vencido", "proximos" → "proximo"
    const estadoNormalizado = estado === 'vencidos' ? 'vencido' : (estado === 'proximos' ? 'proximo' : estado);

    // Extraer clientes seleccionados usando data-attributes (más robusto que split)
    const clientesSeleccionados = Array.from(checkboxes).map(cb => ({
        cliente: cb.getAttribute('data-cliente'),
        email: cb.getAttribute('data-email')
    }));

    console.log(`\n📧 ===== INICIO ENVÍO CORREOS ${estado.toUpperCase()} =====`);
    console.log(`Estado recibido: "${estado}" → Estado normalizado: "${estadoNormalizado}"`);
    console.log(`Checkboxes seleccionados: ${checkboxes.length}`);
    console.log(`Clientes seleccionados:`, clientesSeleccionados);

    if (clientesSeleccionados.length === 0) {
        alert(`No hay clientes seleccionados para ${estado === 'vencidos' ? 'vencidos' : 'próximos'}`);
        return;
    }

    // Filtrar recordatorios por cliente Y email usando comparación normalizada
    const recordatoriosFiltrados = recordatoriosGlobal.filter(r => {
        if (r.estado !== estadoNormalizado) return false;

        const match = clientesSeleccionados.some(cs => {
            const clienteMatch = normalizarTexto(cs.cliente) === normalizarTexto(r.cliente);
            const emailMatch = normalizarTexto(cs.email) === normalizarTexto(r.correo_cliente);
            return clienteMatch && emailMatch;
        });

        return match;
    });

    console.log(`\n🔍 DEBUG FILTRADO:`);
    console.log(`  - Total recordatorios globales: ${recordatoriosGlobal.length}`);
    console.log(`  - Recordatorios con estado "${estadoNormalizado}": ${recordatoriosGlobal.filter(r => r.estado === estadoNormalizado).length}`);
    console.log(`  - Recordatorios filtrados para envío: ${recordatoriosFiltrados.length}`);

    if (recordatoriosFiltrados.length === 0) {
        console.error(`❌ ERROR: No se encontraron recordatorios para los clientes seleccionados`);
        console.log(`\n🔍 DIAGNÓSTICO - Comparando cada cliente seleccionado:`);

        clientesSeleccionados.forEach(cs => {
            const recordatoriosCliente = recordatoriosGlobal.filter(r =>
                normalizarTexto(r.cliente) === normalizarTexto(cs.cliente) &&
                normalizarTexto(r.correo_cliente) === normalizarTexto(cs.email)
            );
            const recordatoriosEstado = recordatoriosCliente.filter(r => r.estado === estadoNormalizado);
            console.log(`  Cliente: "${cs.cliente}" | Email: "${cs.email}"`);
            console.log(`    → Recordatorios totales: ${recordatoriosCliente.length}`);
            console.log(`    → Recordatorios con estado "${estadoNormalizado}": ${recordatoriosEstado.length}`);
            console.log(`    → Estados encontrados:`, [...new Set(recordatoriosCliente.map(r => r.estado))]);
        });

        alert("No hay recordatorios para enviar. Revisa la consola para más detalles.");
        return;
    }

    console.log(`✅ Recordatorios listos para enviar:`, recordatoriosFiltrados);

    const confirmacion = confirm(`¿Enviar ${clientesSeleccionados.length} correos de ${estado === 'vencidos' ? 'facturas vencidas' : 'facturas próximas'}?`);
    if (!confirmacion) {
        console.log(`⚠️ Envío cancelado por el usuario`);
        return;
    }

    const btn = document.getElementById(`btnEnviar${estado.charAt(0).toUpperCase() + estado.slice(1)}`);
    btn.disabled = true;
    btn.textContent = "Enviando...";

    const progressArea = document.getElementById(`progressArea${estado.charAt(0).toUpperCase() + estado.slice(1)}`);
    progressArea.style.display = "block";

    const progressFill = document.getElementById(`progressFill${estado.charAt(0).toUpperCase() + estado.slice(1)}`);
    const progressText = document.getElementById(`progressText${estado.charAt(0).toUpperCase() + estado.slice(1)}`);

    try {
        progressFill.style.width = "20%";
        progressText.textContent = `Enviando ${clientesSeleccionados.length} correos...`;

        console.log(`\n📤 Enviando ${recordatoriosFiltrados.length} recordatorios al servidor...`);

        const response = await fetch("/enviar-correos", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ recordatorios: recordatoriosFiltrados })
        });

        if (!response.ok) {
            const errorData = await response.json();
            console.error(`❌ Error del servidor:`, errorData);
            throw new Error(errorData.message || "Error en servidor");
        }

        const resultado = await response.json();

        console.log(`\n✅ RESULTADO DEL ENVÍO:`);
        console.log(`  - Total correos enviados: ${resultado.total}`);
        console.log(`  - Exitosos: ${resultado.exitosos}`);
        console.log(`  - Fallidos: ${resultado.fallidos}`);
        console.log(`  - Detalles:`, resultado.resultados);
        console.log(`\n===== FIN ENVÍO CORREOS ${estado.toUpperCase()} =====\n`);

        progressFill.style.width = "100%";
        progressText.textContent = "✅ Envío completado";

        document.getElementById(`resultExitosos${estado.charAt(0).toUpperCase() + estado.slice(1)}`).textContent = resultado.exitosos;
        document.getElementById(`resultFallidos${estado.charAt(0).toUpperCase() + estado.slice(1)}`).textContent = resultado.fallidos;

        document.getElementById(`resultsArea${estado.charAt(0).toUpperCase() + estado.slice(1)}`).style.display = "block";

        setTimeout(() => progressArea.style.display = "none", 2000);

    } catch (error) {
        console.error(`❌ ERROR EN ENVÍO:`, error);
        alert("Error: " + error.message);
        progressText.textContent = "❌ Error en envío";
    } finally {
        btn.disabled = false;
        btn.textContent = `Enviar Correos ${estado.charAt(0).toUpperCase() + estado.slice(1)}`;
    }
}
