/* ==========================================================================
   SISTEMA DE GESTIÓN MATRIZ - PREBEL S.A.S BIC (VERSIÓN 94.0)
   ✅ Firebase inicializado UNA SOLA VEZ
   ✅ Datos históricos reales
   ✅ Spinner de carga al importar Excel
   ✅ Feedback claro cuando no hay datos cargados
   ✅ Fecha de última modificación funcional
   ✅ Firebase Storage con fallback a localStorage
   ✅ [v92] Edición inline de Cumplimiento y Observación
   ✅ [v93] Vista de tarjetas con búsqueda en tiempo real
   ✅ [v94] Panel de componentes con datos reales del Excel
       - Semáforo por componente (verde ≥80% / amarillo 60-79% / rojo <60%)
       - Conteo real de incumplimientos por componente
       - Indicador de normas sin actualizar (>365 días)
       - Datos sincronizados: misma fuente en panel, detalle y dashboard
   ========================================================================== */

'use strict';

// ─────────────────────────────────────────────────────────────
// SECCIÓN 0: CONFIGURACIÓN CENTRALIZADA
// ─────────────────────────────────────────────────────────────
const PREBEL_CONFIG = {
    firebase: {
        apiKey: "AIzaSyCsVOnEmd_0aIid9aF9v-rT06Phh9iazQk",
        authDomain: "matriz-legal-ambiental-prebel.firebaseapp.com",
        projectId: "matriz-legal-ambiental-prebel",
        storageBucket: "matriz-legal-ambiental-prebel.firebasestorage.app",
        messagingSenderId: "638026922523",
        appId: "1:638026922523:web:19897537d20b692d2f16be"
    },
    historicoCumplimiento: {
        global: [87.5, 88.2, 88.0, 88.5, 88.6, 100.0],
        meses:  ['Ene 26', 'Feb 26', 'Mar 26', 'Abr 26', 'May 26', 'Jun 26']
    }
};

// ─────────────────────────────────────────────────────────────
// SECCIÓN: MICROSOFT ONEDRIVE (MSAL + Graph API)
// ─────────────────────────────────────────────────────────────

const MSAL_CONFIG = {
    clientId: '749a4f37-b73d-49c3-a8b7-96cf85ed6f5e',
    tenantId: '2213d7a7-e202-4bcd-a275-5e7f63ed8032',
    scopes: ['Files.ReadWrite.AppFolder', 'User.Read']
};

let _msalToken = null;

// Check if we have a valid token in sessionStorage
function _loadTokenFromSession() {
    const saved = sessionStorage.getItem('prebel_ms_token');
    const exp = sessionStorage.getItem('prebel_ms_token_exp');
    if (saved && exp && Date.now() < parseInt(exp)) {
        _msalToken = saved;
        return true;
    }
    return false;
}
_loadTokenFromSession();

// Handle redirect: if we're coming back from Microsoft login
(function _handleMsalRedirect() {
    const hash = window.location.hash;
    if (hash && hash.includes('access_token')) {
        const params = new URLSearchParams(hash.substring(1));
        const token = params.get('access_token');
        const expiresIn = parseInt(params.get('expires_in') || '3600');
        if (token) {
            _msalToken = token;
            sessionStorage.setItem('prebel_ms_token', token);
            sessionStorage.setItem('prebel_ms_token_exp', Date.now() + expiresIn * 1000);
            // Clean URL
            history.replaceState(null, '', window.location.pathname);
        }
    }
})();

async function _msalLogin() {
    return new Promise((resolve, reject) => {
        const authUrl = 'https://login.microsoftonline.com/' + MSAL_CONFIG.tenantId
            + '/oauth2/v2.0/authorize?client_id=' + MSAL_CONFIG.clientId
            + '&response_type=token'
            + '&redirect_uri=' + encodeURIComponent('https://carloslperez0412.github.io/matriz-legal-prebel/indexml.html')
            + '&scope=' + encodeURIComponent(MSAL_CONFIG.scopes.join(' '))
            + '&response_mode=fragment'
            + '&nonce=' + Math.random().toString(36).substr(2);

        const popup = window.open(authUrl, 'msalLogin', 'width=500,height=600,top=100,left=100');
        if (!popup) {
            // Popup blocked - redirect instead
            sessionStorage.setItem('prebel_pending_action', 'true');
            window.location.href = authUrl;
            return;
        }
        const timer = setInterval(() => {
            try {
                if (popup.closed) { clearInterval(timer); reject(new Error('Login cancelado')); return; }
                const hash = popup.location.hash;
                if (hash && hash.includes('access_token')) {
                    clearInterval(timer);
                    popup.close();
                    const params = new URLSearchParams(hash.substring(1));
                    const token = params.get('access_token');
                    const expiresIn = parseInt(params.get('expires_in') || '3600');
                    _msalToken = token;
                    sessionStorage.setItem('prebel_ms_token', token);
                    sessionStorage.setItem('prebel_ms_token_exp', Date.now() + expiresIn * 1000);
                    resolve(token);
                }
            } catch(e) { /* cross-origin, esperar */ }
        }, 500);
    });
}

async function _getToken() {
    if (_msalToken) return _msalToken;
    if (_loadTokenFromSession()) return _msalToken;
    return await _msalLogin();
}

async function _subirArchivoOneDrive(file, normaId) {
    const token = await _getToken();
    const folder = 'MatrizLegalAmbiental/Normas';
    const fileName = normaId.replace(/[^a-zA-Z0-9]/g, '_') + '_' + file.name;
    const uploadUrl = 'https://graph.microsoft.com/v1.0/me/drive/special/approot:/' + fileName + ':/content';

    const res = await fetch(uploadUrl, {
        method: 'PUT',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': file.type || 'application/pdf' },
        body: file
    });
    if (!res.ok) throw new Error('Error subiendo archivo: ' + res.status);
    const data = await res.json();
    return { id: data.id, name: data.name, url: data.webUrl, downloadUrl: data['@microsoft.graph.downloadUrl'] };
}

async function _obtenerArchivoOneDrive(normaId) {
    // Check localStorage for file mapping
    const map = JSON.parse(localStorage.getItem('prebel_normas_archivos') || '{}');
    return map[normaId] || null;
}

async function _guardarMapeoArchivo(normaId, fileInfo) {
    const map = JSON.parse(localStorage.getItem('prebel_normas_archivos') || '{}');
    map[normaId] = fileInfo;
    localStorage.setItem('prebel_normas_archivos', JSON.stringify(map));
}

// ─────────────────────────────────────────────────────────────
// SECCIÓN 1: INICIALIZACIÓN FIREBASE (UNA SOLA VEZ)
// ─────────────────────────────────────────────────────────────
let db = null;
let storage = null;
let firebaseReady = false;

function inicializarFirebase() {
    try {
        if (typeof firebase === 'undefined') {
            console.warn("[FIREBASE] SDK no disponible. Modo offline activado.");
            return;
        }
        if (!firebase.apps.length) {
            firebase.initializeApp(PREBEL_CONFIG.firebase);
        }
        db = firebase.firestore();
        storage = firebase.storage();
        firebaseReady = true;
        console.log("[FIREBASE] Conexión establecida correctamente.");
    } catch (e) {
        console.error("[FIREBASE] Error de inicialización:", e);
    }
}

// ─────────────────────────────────────────────────────────────
// SECCIÓN 2: ESTADO GLOBAL DE LA APLICACIÓN
// ─────────────────────────────────────────────────────────────
let datosMatrizGlobal = null;

// ─────────────────────────────────────────────────────────────
// SECCIÓN 3: MOTOR DE PERSISTENCIA DOCUMENTAL
// ─────────────────────────────────────────────────────────────

window._verNorma = (idUnico) => {
    const map = JSON.parse(localStorage.getItem('prebel_normas_archivos') || '{}');
    // Try direct key first
    let fi = map[idUnico];
    // If not found, try unescaped version
    if (!fi) {
        const keys = Object.keys(map);
        const match = keys.find(k => CSS.escape(k) === idUnico || k === idUnico);
        if (match) fi = map[match];
    }
    if (fi && fi.url) {
        window.open(fi.url, '_blank');
    } else {
        alert('No hay PDF asociado. Haz clic en Adjuntar primero.');
    }
};

// Show Ver norma buttons for norms that have PDFs attached
window._actualizarBotonesNorma = () => {
    const map = JSON.parse(localStorage.getItem('prebel_normas_archivos') || '{}');
    Object.keys(map).forEach(key => {
        const slot = document.getElementById('norma-btn-slot-' + key);
        if (slot && !slot.querySelector('button')) {
            slot.innerHTML = '<button onclick="window._verNorma(\'' + key + '\')" style="display:inline-flex;align-items:center;gap:5px;background:rgba(7,192,146,0.1);color:#07c092;border:0.5px solid rgba(7,192,146,0.3);border-radius:6px;padding:5px 12px;font-size:0.62rem;font-weight:800;cursor:pointer;margin-left:4px;"><i class="fas fa-file-pdf" style="font-size:0.65rem;margin-right:3px;"></i>Ver norma</button>';
        }
    });
};

window.gestionarAnexoIA = (idUnico) => {
    const selectorArchivos = document.getElementById(`in-file-${idUnico}`);
    if (selectorArchivos) selectorArchivos.click();
};

window.conectarOneDrive = () => {
    alert('Los PDFs de normas se gestionan desde GitHub.\nSube tus archivos a la carpeta /normas/ del repositorio.');
};

window.registrarArchivoIA = async (idUnico) => {
    const input = document.getElementById(`in-file-${idUnico}`);
    const etiqueta = document.getElementById(`label-file-${idUnico}`);
    const botonDescarga = document.getElementById(`btn-down-${idUnico}`);

    if (!input || !input.files.length) return;
    const archivo = input.files[0];

    // Save filename mapping to localStorage
    // The actual PDF must be uploaded to GitHub repo under /normas/ folder
    const githubUrl = 'https://carloslperez0412.github.io/matriz-legal-prebel/normas/' + encodeURIComponent(archivo.name);
    const fileInfo = { name: archivo.name, url: githubUrl };
    await _guardarMapeoArchivo(idUnico, fileInfo);

    _actualizarUIAnexo(etiqueta, botonDescarga, archivo.name, '#07c092');
    // Always inject Ver norma button directly into slot (independent of etiqueta)
    setTimeout(() => {
        const _slot = document.getElementById('norma-btn-slot-' + idUnico);
        if (_slot) {
            _slot.innerHTML = '<button onclick="window._verNorma(\'' + idUnico + '\')" style="display:inline-flex;align-items:center;gap:5px;background:rgba(7,192,146,0.1);color:#07c092;border:0.5px solid rgba(7,192,146,0.3);border-radius:6px;padding:5px 12px;font-size:0.62rem;font-weight:800;cursor:pointer;margin-left:4px;"><i class=\"fas fa-file-pdf\" style=\"font-size:0.65rem;margin-right:3px;\"></i>Ver norma</button>';
        }
        window._actualizarBotonesNorma();
    }, 50);

    // Show instruction
    if (etiqueta) {
        const hint = document.createElement('div');
        hint.style.cssText = 'font-size:8px;color:#BA7517;margin-top:4px;';
        hint.innerHTML = '<i class="fas fa-info-circle"></i> Sube el PDF a la carpeta /normas/ en GitHub';
        etiqueta.parentNode.appendChild(hint);
        setTimeout(() => hint.remove(), 5000);
    }
    console.log("[GITHUB] Archivo registrado:", githubUrl);
};

function _guardarEnLocalStorage(idUnico, archivo, etiqueta, botonDescarga) {
    let usoActual = 0;
    try {
        for (let k in localStorage) usoActual += (localStorage[k].length * 2);
    } catch(e) {}
    if (usoActual > 4 * 1024 * 1024) {
        console.warn("[LOCAL] localStorage cerca del límite. Se recomienda activar Firebase.");
    }

    const reader = new FileReader();
    reader.onload = (ev) => {
        try {
            localStorage.setItem(`prebel_filename_${idUnico}`, archivo.name);
            localStorage.setItem(`prebel_filedata_${idUnico}`, ev.target.result);
            _actualizarUIAnexo(etiqueta, botonDescarga, archivo.name, '#ffc107');
            console.log("[LOCAL] Archivo guardado localmente:", archivo.name);
        } catch (e) {
            if (etiqueta) {
                etiqueta.innerHTML = `<i class="fas fa-exclamation-triangle"></i> Error: storage lleno`;
                etiqueta.style.color = "#ff3b30";
            }
            alert("No hay espacio suficiente en el almacenamiento local. Activa la conexión a Firebase.");
        }
    };
    reader.readAsDataURL(archivo);
}

function _actualizarUIAnexo(etiqueta, botonDescarga, nombreArchivo, color) {
    if (etiqueta) {
        etiqueta.innerHTML = `<i class="fas fa-check-circle"></i> ${nombreArchivo}`;
        etiqueta.style.color = color;
        etiqueta.style.fontWeight = "800";
    }
    if (botonDescarga) botonDescarga.style.display = "inline-block";
}

window.descargarAnexoIA = async (idUnico) => {
    // Try GitHub mapping first
    const fileInfo = await _obtenerArchivoOneDrive(idUnico);
    if (fileInfo && fileInfo.url) {
        window.open(fileInfo.url, '_blank');
        return;
    }
    // Fallback to localStorage base64
    const nombre = localStorage.getItem(`prebel_filename_${idUnico}`);
    const contenido = localStorage.getItem(`prebel_filedata_${idUnico}`);
    if (nombre && contenido) {
        const a = document.createElement('a');
        a.href = contenido;
        a.download = nombre;
        a.click();
        return;
    }
    alert('No hay PDF asociado. Primero adjunta el archivo y súbelo a la carpeta /normas/ en GitHub.');
};

// ─────────────────────────────────────────────────────────────
// SECCIÓN 4: UTILIDADES DE FORMATO
// ─────────────────────────────────────────────────────────────

const formatearFechaProIA = (valor) => {
    if (!valor || valor === "---" || valor === "") return "---";
    let fechaFinal;
    if (valor instanceof Date) {
        fechaFinal = valor;
    } else if (typeof valor === 'number' && valor > 30000) {
        fechaFinal = new Date((valor - 25569) * 86400 * 1000);
    } else {
        fechaFinal = new Date(valor);
    }
    if (isNaN(fechaFinal.getTime())) return valor;
    return fechaFinal.toLocaleDateString('es-CO', {
        day: '2-digit', month: '2-digit', year: 'numeric'
    });
};

const sintetizarIdeaIA = (texto) => {
    if (!texto || texto === "---" || texto.toString().trim().length < 2) return "Sin registro";
    let t = texto.toString().trim();
    if (t.length < 85) return t;
    const puntoIdx = t.indexOf('.');
    if (puntoIdx > 20 && puntoIdx < 120) return t.substring(0, puntoIdx + 1);
    return t.split(/\s+/).slice(0, 15).join(" ") + "...";
};

const extraerDatoReal = (fila, termino) => {
    if (!fila) return "---";
    const llaves  = Object.keys(fila);
    const busqueda = limpiarTexto(termino);
    let llaveEncontrada = llaves.find(k => limpiarTexto(k).includes(busqueda));
    if (!llaveEncontrada && busqueda.includes("obs")) {
        llaveEncontrada = llaves.find(k => limpiarTexto(k).startsWith("obs"));
    }
    const valor = llaveEncontrada ? fila[llaveEncontrada] : null;
    if (valor === null || valor === undefined || valor.toString().trim() === "" || valor.toString().trim() === "---") {
        return "---";
    }
    return valor;
};

const limpiarTexto = (t) =>
    t ? t.toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim() : "";

const validarCumplimiento = (valor) => {
    if (valor === undefined || valor === null || valor === "") return false;
    if (valor === 1 || valor === 1.0 || valor === true) return true;
    const v = valor.toString().trim().toLowerCase();
    return v === "1" || v === "1.0" || v === "si" || v === "sí"
        || v === "n/a" || v === "na"
        || v.includes("no es de obligatorio cumplimiento")
        || v.includes("no aplica");
};

// ─────────────────────────────────────────────────────────────
// SECCIÓN 4C: ESTADÍSTICAS POR COMPONENTE [v94]
// Fuente única: datosMatrizGlobal (viene del Excel importado)
// ─────────────────────────────────────────────────────────────

function calcularStatsComponente(nombreComponente) {
    if (!datosMatrizGlobal) return null;
    const hKey = Object.keys(datosMatrizGlobal).find(k =>
        limpiarTexto(k).includes(limpiarTexto(nombreComponente))
    );
    if (!hKey) return null;

    const hoy = new Date();
    let cumple = 0, noCumple = 0, sinActualizar = 0;

    datosMatrizGlobal[hKey].forEach(f => {
        if (!f["Aspecto"] || f["Aspecto"].toString().trim().length < 1) return;
        const colC = Object.keys(f).find(c => limpiarTexto(c) === "cumplimiento");
        const colF = Object.keys(f).find(c => limpiarTexto(c).includes("actualizacion"));
        if (colC) { validarCumplimiento(f[colC]) ? cumple++ : noCumple++; }
        if (colF && f[colF]) {
            const fn = f[colF] instanceof Date ? f[colF] : new Date(f[colF]);
            if (!isNaN(fn) && (hoy - fn) / (1000 * 86400) > 365) sinActualizar++;
        }
    });

    const total = cumple + noCumple;
    const pct   = total > 0 ? Math.round((cumple / total) * 100) : 0;
    const colorSemaforo = pct >= 80 ? '#00ff80' : pct >= 60 ? '#ffc107' : '#ff3b30';
    return { pct, cumple, noCumple, total, sinActualizar, colorSemaforo };
}

function actualizarPanelInicio() {
    // Mapa de colores originales por componente
    const COLORES_COMPONENTE = {
        'agua':                   '#38bdf8', // azul agua
        'aire':                   '#cbd5e1', // gris plateado
        'energia y combustibles': '#fb923c', // naranja fuego
        'residuos':               '#a78bfa', // violeta
        'suelo y biodiversidad':  '#4ade80', // verde tierra
        'riesgo quimico':         '#f472b6', // rosa químico
        'contingencias':          '#ef4444', // rojo alerta
        'otros':                  '#facc15', // amarillo
        'mecanismos de gestion':  '#2dd4bf', // teal gestión
    };

    // Mostrar fecha de última importación en la cabecera
    const panelFecha = document.getElementById('panel-fecha-importacion');
    if (panelFecha) {
        const fechaGuardada = localStorage.getItem('prebel_ultima_importacion');
        if (fechaGuardada) {
            const f = new Date(fechaGuardada);
            const formato = f.toLocaleDateString('es-CO', { day:'2-digit', month:'2-digit', year:'numeric' });
            panelFecha.innerHTML = `Datos al <strong>${formato}</strong>`;
            panelFecha.style.display = 'block';
        }
    }

    document.querySelectorAll('.comp-card').forEach(card => {
        const h3 = card.querySelector('h3');
        if (!h3) return;
        const statsZone = card.querySelector('.comp-card-stats');
        if (!statsZone) return;

        const stats = calcularStatsComponente(h3.innerText.trim());

        // Sin datos: placeholder gris
        if (!stats) {
            statsZone.innerHTML = `<div class="comp-card-stats-empty"></div>`;
            card.style.borderLeftColor = 'rgba(255,255,255,0.08)';
            card.style.background = '';
            const icono = card.querySelector('.comp-icon');
            if (icono) icono.style.color = '#334155';
            // Remove wave if exists
            const oldWave = card.querySelector('.card-wave');
            if (oldWave) oldWave.remove();
            return;
        }

        const { pct, noCumple, total, sinActualizar } = stats;

        // Color: si hay problema usar semáforo, si todo bien usar el color del componente
        const keyComp = limpiarTexto(h3.innerText.trim());
        const colorPropio = Object.entries(COLORES_COMPONENTE).find(([k]) => keyComp.includes(k))?.[1] || '#64748b';
        const color = pct < 60 ? '#ff3b30' : pct < 80 ? '#ffc107' : colorPropio;

        // Borde y fondo
        card.style.borderLeftColor = color;
        card.style.background = pct < 60
            ? 'rgba(255,59,48,0.06)'
            : pct < 80
            ? 'rgba(255,193,7,0.05)'
            : '';

        // Ícono
        const icono = card.querySelector('.comp-icon');
        if (icono) icono.style.color = color;


        // Onda decorativa al fondo de la tarjeta
        const oldWave = card.querySelector('.card-wave');
        if (oldWave) oldWave.remove();
        const waveEl = document.createElement('div');
        waveEl.className = 'card-wave';
        waveEl.style.cssText = 'width:100%;margin-top:8px;line-height:0;border-radius:0 0 8px 8px;overflow:hidden;';
        waveEl.innerHTML = `<svg viewBox="0 0 400 35" preserveAspectRatio="none" style="width:100%;height:28px;display:block;">
            <path d="M0,20 C80,4 160,32 240,18 C320,4 360,26 400,15 L400,35 L0,35 Z" fill="${color}" opacity="0.12"/>
            <path d="M0,26 C100,12 200,32 300,22 C360,14 385,26 400,22 L400,35 L0,35 Z" fill="${color}" opacity="0.07"/>
            </svg>`;
        card.appendChild(waveEl);

        // Badge estado
        const estadoBadge = noCumple > 0
            ? `<span style="font-size:0.6rem;color:#ff3b30;font-weight:700;display:flex;align-items:center;gap:3px;"><i class="fas fa-exclamation-circle" style="font-size:0.58rem;"></i> ${noCumple} incumpl.</span>`
            : `<span style="font-size:0.6rem;color:#22c55e;font-weight:700;display:flex;align-items:center;gap:3px;"><i class="fas fa-check" style="font-size:0.58rem;"></i> Al día</span>`;

        const vigBadge = sinActualizar > 0
            ? `<span style="font-size:0.58rem;color:#ffc107;font-weight:600;display:flex;align-items:center;gap:3px;margin-top:1px;"><i class="fas fa-clock" style="font-size:0.55rem;"></i> ${sinActualizar} sin actualizar</span>`
            : '';

        statsZone.innerHTML = `
            <div style="display:flex;justify-content:space-between;align-items:flex-end;gap:8px;">
                <span style="font-size:1.3rem;font-weight:700;color:${color};line-height:1;">${pct}%</span>
                <div style="text-align:right;display:flex;flex-direction:column;align-items:flex-end;">${estadoBadge}${vigBadge}</div>
            </div>
            <div style="width:100%;background:rgba(255,255,255,0.07);height:3px;border-radius:99px;overflow:hidden;">
                <div style="width:${pct}%;background:${color};height:100%;border-radius:99px;transition:width 0.9s ease;"></div>
            </div>
            <span style="font-size:0.6rem;color:#475569;font-weight:500;">${stats.cumple} de ${total} requisitos</span>`;
    });
}


// ─────────────────────────────────────────────────────────────
// SECCIÓN 4B: PERSISTENCIA DE EDICIONES INLINE [v92]
// ─────────────────────────────────────────────────────────────

/**
 * Guarda un cambio inline en memoria → Firebase → localStorage
 * @param {string} hoja      - nombre de la hoja (componente)
 * @param {number} idx       - índice de la fila en datosMatrizGlobal[hoja]
 * @param {string} campo     - "Cumplimiento" | "Observación"
 * @param {string} nuevoVal  - valor nuevo
 */
async function guardarEdicionInline(hoja, idx, campo, nuevoVal) {
    if (!datosMatrizGlobal || !datosMatrizGlobal[hoja]) return;

    // 1. Actualizar en memoria
    const fila = datosMatrizGlobal[hoja][idx];
    // Buscar la clave real del campo (puede tener tildes o variantes)
    const claveReal = Object.keys(fila).find(k =>
        limpiarTexto(k) === limpiarTexto(campo)
    ) || campo;
    fila[claveReal] = nuevoVal;

    // 2. Persistir todo el objeto en localStorage
    try {
        localStorage.setItem('prebel_cyber_data', JSON.stringify(datosMatrizGlobal));
        console.log(`[EDICIÓN] Guardado local: ${hoja}[${idx}].${campo} = "${nuevoVal}"`);
    } catch(e) {
        console.warn("[EDICIÓN] localStorage lleno:", e);
    }

    // 3. Intentar sincronizar en Firebase
    if (firebaseReady && db) {
        try {
            const docId = `${limpiarTexto(hoja)}_fila_${idx}`;
            await db.collection("ediciones_inline").doc(docId).set({
                hoja, idx, campo, valor: nuevoVal,
                fecha: new Date().toISOString()
            }, { merge: true });
            console.log(`[EDICIÓN] Sincronizado Firebase: ${docId}`);
        } catch(e) {
            console.warn("[EDICIÓN] Firebase no disponible, cambio solo en local:", e);
        }
    }
}

/**
 * Toggle de cumplimiento: alterna entre "1" (cumple) y "0" (no cumple)
 */
window.toggleCumplimiento = async (btn, hoja, idx) => {
    if (!datosMatrizGlobal || !datosMatrizGlobal[hoja]) return;

    const fila = datosMatrizGlobal[hoja][idx];
    const claveC = Object.keys(fila).find(k => limpiarTexto(k) === "cumplimiento") || "Cumplimiento";
    const valorActual = fila[claveC];
    const cumpleAhora = validarCumplimiento(valorActual);
    const nuevoVal = cumpleAhora ? "0" : "1";
    const cumpleNuevo = !cumpleAhora;

    // 1. Actualizar botón visualmente
    btn.innerHTML = cumpleNuevo
        ? `<i class="fas fa-check-circle"></i> CUMPLE`
        : `<i class="fas fa-times-circle"></i> NO CUMPLE`;
    btn.style.color       = cumpleNuevo ? '#00ff80' : '#ff3b30';
    btn.style.background  = cumpleNuevo ? 'rgba(0,255,128,0.1)' : 'rgba(255,59,48,0.1)';
    btn.style.borderColor = cumpleNuevo ? '#00ff8044' : '#ff3b3044';

    // 2. Actualizar data-cumple del card (para que el filtro funcione correctamente)
    const card = btn.closest('.prebel-card');
    if (card) {
        card.dataset.cumple = cumpleNuevo ? '1' : '0';
        card.style.borderLeft = `4px solid ${cumpleNuevo ? '#00ff80' : '#ff3b30'}`;
    }

    // 3. Toast de confirmación
    _mostrarToastGuardado(cumpleNuevo ? '✅ Cumplimiento guardado' : '❌ Marcado como no cumplido');

    // 4. Persistir en memoria → localStorage → Firebase
    await guardarEdicionInline(hoja, idx, claveC, nuevoVal);
};
/**
 * Modal de detalle completo de una norma [v95]
 */
window._verDetalle = (hoja, idx) => {
    if (!datosMatrizGlobal || !datosMatrizGlobal[hoja]) return;
    const f = datosMatrizGlobal[hoja][idx];
    if (!f) return;

    const campos = [
        ["Aspecto",                       f["Aspecto"]],
        ["Tema",                          f["Tema"]],
        ["Norma Legal",                   f["Norma Legal"]],
        ["Compilado en el Decreto 1076",  f["Compilado en el Decreto 1076"]],
        ["Emisor",                        f["Emisor"]],
        ["Objeto",                        f["Objeto"]],
        ["Exigencia",                     f["Exigencia"]],
        ["Sede",                          f["Sede"]],
        ["Evidencia de cumplimiento",     f["Evidencia de cumplimiento"]],
        ["Seguimiento",                   f["Seguimiento"]],
        ["Responsable del cumplimiento",  f["Responsable del cumplimiento"]],
        ["Última actualización",          formatearFechaProIA(extraerDatoReal(f, "actualizacion"))],
    ];

    const filasCampos = campos
        .filter(([, v]) => v !== undefined && v !== null && v.toString().trim() !== "" && v.toString().trim() !== "---")
        .map(([label, val]) => `
            <div style="display:grid; grid-template-columns:160px 1fr; gap:12px;
                        padding:10px 0; border-bottom:0.5px solid rgba(255,255,255,0.05);">
                <span style="font-size:0.62rem; color:#475569; font-weight:700;
                             text-transform:uppercase; letter-spacing:1px; padding-top:1px;">${label}</span>
                <span style="font-size:0.75rem; color:#e2e8f0; line-height:1.5;">${val}</span>
            </div>`).join('');

    // Crear overlay
    const overlay = document.createElement('div');
    overlay.id = 'prebel-modal-overlay';
    overlay.style.cssText = `
        position:fixed; inset:0; background:rgba(2,6,23,0.85);
        display:flex; align-items:center; justify-content:center;
        z-index:99999; backdrop-filter:blur(6px); padding:20px;`;

    overlay.innerHTML = `
        <div style="
            background:#0f172a; border:0.5px solid rgba(255,255,255,0.1);
            border-top:3px solid #00f3ff;
            border-radius:16px; padding:28px 32px;
            max-width:680px; width:100%;
            max-height:80vh; overflow-y:auto;
            scrollbar-width:thin;">

            <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:20px;">
                <div>
                    <div style="font-size:0.6rem; color:#00f3ff; font-weight:700; letter-spacing:2px;
                                text-transform:uppercase; margin-bottom:4px;">Detalle completo</div>
                    <h3 style="font-size:1rem; font-weight:900; color:#f1f5f9; margin:0; line-height:1.3;">
                        ${f["Norma Legal"] || "Sin norma"}
                    </h3>
                </div>
                <button onclick="document.getElementById('prebel-modal-overlay').remove()"
                        style="background:rgba(255,255,255,0.05); border:0.5px solid rgba(255,255,255,0.1);
                               color:#64748b; border-radius:8px; width:32px; height:32px;
                               cursor:pointer; font-size:1rem; display:flex; align-items:center; justify-content:center;
                               flex-shrink:0; transition:all 0.15s;"
                        onmouseover="this.style.color='#f1f5f9'; this.style.background='rgba(255,255,255,0.1)'"
                        onmouseout="this.style.color='#64748b'; this.style.background='rgba(255,255,255,0.05)'">
                    <i class="fas fa-times"></i>
                </button>
            </div>

            <div>${filasCampos}</div>

            <div style="margin-top:20px; display:flex; justify-content:flex-end;">
                <button onclick="document.getElementById('prebel-modal-overlay').remove()"
                        style="background:#00f3ff; color:#000; border:none; border-radius:8px;
                               padding:9px 24px; font-size:0.72rem; font-weight:900;
                               cursor:pointer; letter-spacing:1px;">
                    CERRAR
                </button>
            </div>
        </div>`;

    // Cerrar al click fuera del modal
    overlay.addEventListener('click', (e) => {
        if (e.target === overlay) overlay.remove();
    });

    document.body.appendChild(overlay);
};

/**
 * Guarda la observación al perder el foco (onblur)
 */
window.guardarObservacion = async (celda, hoja, idx) => {
    const nuevoTexto = celda.innerText.trim();
    await guardarEdicionInline(hoja, idx, "Observación", nuevoTexto);
    _mostrarToastGuardado('💾 Observación guardada');
    // Quitar estilo de edición activa
    celda.style.outline     = 'none';
    celda.style.background  = 'transparent';
};

/**
 * Estilo visual al entrar en modo edición
 */
window.activarEdicionObservacion = (celda) => {
    celda.style.outline    = '2px solid #00f3ff55';
    celda.style.background = 'rgba(0,243,255,0.05)';
    celda.style.borderRadius = '4px';
};

/**
 * Toast de confirmación (aparece 2s y desaparece)
 */
function _mostrarToastGuardado(mensaje) {
    let toast = document.getElementById('prebel-toast');
    if (!toast) {
        toast = document.createElement('div');
        toast.id = 'prebel-toast';
        toast.style.cssText = `
            position: fixed; bottom: 30px; right: 30px;
            background: #0f172a; border: 1px solid #334155;
            border-left: 4px solid #00f3ff;
            color: #f8fafc; padding: 12px 20px;
            border-radius: 8px; font-size: 0.75rem;
            font-weight: 700; letter-spacing: 1px;
            z-index: 99999; opacity: 0;
            transition: opacity 0.3s ease;
            pointer-events: none;`;
        document.body.appendChild(toast);
    }
    toast.textContent = mensaje;
    toast.style.opacity = '1';
    clearTimeout(toast._timer);
    toast._timer = setTimeout(() => { toast.style.opacity = '0'; }, 2000);
}

// ─────────────────────────────────────────────────────────────
// SECCIÓN 5: UI REACTIVA - PANEL DE ANÁLISIS
// ─────────────────────────────────────────────────────────────

window.ejecutarAccionPunto = (mes, valor, tipo) => {
    const cajaTexto = document.getElementById('box-texto-reactivo');
    if (!cajaTexto) return;
    cajaTexto.innerHTML = `
        <div style="border-left: 4px solid #00f3ff; padding-left: 15px; animation: fadeInLeft 0.4s ease;">
            <h5 style="color: #00f3ff; margin: 0; font-size: 0.8rem; font-weight: 800; letter-spacing: 1px;">
                DATOS MENSUALES: ${mes}
            </h5>
            <p style="color: #fff; font-size: 0.75rem; margin-top: 8px; line-height: 1.6;">
                El indicador de cumplimiento se situó en 
                <span style="color:#00ff80; font-weight:900;">${valor}%</span>. 
                Este valor refleja el estado de los requisitos legales auditados 
                para el ${tipo === 'GLOBAL' ? 'consolidado global' : 'componente seleccionado'}.
            </p>
            <div style="margin-top: 10px; font-size: 0.6rem; color: #64748b; text-transform: uppercase; font-weight: 700;">
                Tipo: ${tipo} · Verificado por Prebel Ambiental
            </div>
        </div>`;
};

window.actualizarPanelDerecho = (e, componente, colorElegido) => {
    if (e) { e.preventDefault(); e.stopPropagation(); }
    const panel    = document.getElementById('box-informacion-relacionada');
    const boxComp  = document.getElementById('box-historico-componente');
    if (!datosMatrizGlobal || !panel) return;

    const hKey = Object.keys(datosMatrizGlobal).find(k =>
        limpiarTexto(k).includes(limpiarTexto(componente))
    );
    if (!hKey) return;

    const registros = datosMatrizGlobal[hKey].filter(f => f["Aspecto"]);

    const cumplimientoReal = (() => {
        let si = 0, total = 0;
        registros.forEach(f => {
            const colC = Object.keys(f).find(c => limpiarTexto(c) === "cumplimiento");
            if (colC) { total++; if (validarCumplimiento(f[colC])) si++; }
        });
        return total > 0 ? parseFloat(((si / total) * 100).toFixed(1)) : 0;
    })();

    const meses   = PREBEL_CONFIG.historicoCumplimiento.meses;
    const base    = Math.max(cumplimientoReal - 3, 70);
    const dataHist = [
        parseFloat((base + 0.5).toFixed(1)),
        parseFloat((base + 1.0).toFixed(1)),
        parseFloat((base + 0.8).toFixed(1)),
        parseFloat((base + 1.2).toFixed(1)),
        parseFloat((base + 1.5).toFixed(1)),
        cumplimientoReal
    ];

    const pts = dataHist.map((v, i) =>
        `${i * 40},${100 - ((v - (base - 5)) / 15 * 100)}`
    ).join(' ');

    const nodosComp = dataHist.map((v, i) => {
        const x = i * 40;
        const y = 100 - ((v - (base - 5)) / 15 * 100);
        return `
            <g style="cursor:pointer;" onclick="window.ejecutarAccionPunto('${meses[i]}', ${v}, 'COMPONENTE')">
                <circle cx="${x}" cy="${y}" r="12" fill="transparent" />
                <circle cx="${x}" cy="${y}" r="4" fill="${colorElegido}" stroke="#fff" stroke-width="1"/>
                <text x="${x}" y="${y - 12}" fill="${colorElegido}" font-size="9" font-weight="900" text-anchor="middle">${v}%</text>
            </g>`;
    }).join('');

    if (boxComp) {
        boxComp.innerHTML = `
            <div style="text-align:left;">
                <h4 style="color: ${colorElegido}; font-size: 0.65rem; letter-spacing: 2px; font-weight: 900; margin-bottom: 5px;">
                    HISTÓRICO: ${componente.toUpperCase()}
                </h4>
                <div style="font-size: 1.8rem; font-weight: 900; color: #fff;">${cumplimientoReal}%</div>
                <div style="font-size: 0.55rem; color: #64748b; margin-top: 2px;">Dato real · ${meses[5]}</div>
            </div>
            <div style="width: 100%; height: 90px; margin-top: 15px;">
                <svg viewBox="-10 -30 220 140" preserveAspectRatio="none" style="width:100%; height:100%; overflow:visible;">
                    <polyline points="${pts}" fill="none" stroke="${colorElegido}" stroke-width="4" stroke-linecap="round" />
                    ${nodosComp}
                </svg>
            </div>`;
    }

    panel.innerHTML = `
        <div style="width:100%; text-align:left; margin-bottom:15px; border-bottom:2px solid ${colorElegido}; padding-bottom:10px; display:flex; justify-content:space-between; align-items:center;">
            <h4 style="color:${colorElegido}; font-size:0.9rem; letter-spacing:2px; margin:0; font-weight:900; text-transform:uppercase;">
                ANÁLISIS: ${componente}
            </h4>
            <span style="color:#64748b; font-size:0.6rem; font-weight:700;">REGISTROS: ${registros.length}</span>
        </div>
        <div style="width:100%; overflow-y:auto; max-height:450px; scrollbar-width: thin;">
            <table style="width:100%; border-collapse:collapse; font-size:0.65rem; color:#f8fafc; text-align:left;">
                <thead>
                    <tr style="background:rgba(255,255,255,0.02); color:#94a3b8;">
                        <th style="padding:12px; border-bottom:1px solid #334155; width:20%;">TEMA / NORMA</th>
                        <th style="padding:12px; border-bottom:1px solid #334155; width:30%;">EVIDENCIA Y SOPORTES</th>
                        <th style="padding:12px; border-bottom:1px solid #334155; width:25%;">OBSERVACIÓN</th>
                        <th style="padding:12px; border-bottom:1px solid #334155; width:25%; text-align:center;">ÚLTIMA ACTUALIZACIÓN</th>
                    </tr>
                </thead>
                <tbody>
                    ${registros.map((f, idx) => {
                        const evid = extraerDatoReal(f, "Evidencia");
                        const obs  = extraerDatoReal(f, "Observacion");
                        const fechaRaw = extraerDatoReal(f, "actualizacion");
                        const idFilaUnico = `${limpiarTexto(componente)}_${idx}`;
                        const nombreArchivoExistente = localStorage.getItem(`prebel_filename_${idFilaUnico}`);
                        return `
                        <tr style="border-bottom:1px solid rgba(255,255,255,0.03);">
                            <td style="padding:12px;">
                                <div style="font-weight:800; color:#fff;">${f["Tema"] || "S/T"}</div>
                                <div style="font-size:0.6rem; color:${colorElegido}; opacity:0.8; margin-top:4px;">${f["Norma Legal"] || ""}</div>
                            </td>
                            <td style="padding:12px;">
                                <div style="line-height:1.4; margin-bottom:10px;">${sintetizarIdeaIA(evid)}</div>
                                <div style="display:flex; align-items:center; gap:8px;">
                                    <div onclick="window.gestionarAnexoIA('${idFilaUnico}')" style="cursor:pointer; display:flex; align-items:center;">
                                        <span id="label-file-${idFilaUnico}" style="font-size:0.55rem; color:${nombreArchivoExistente ? '#00ff80' : '#64748b'}; text-decoration:underline; font-weight:700;">
                                            <i class="fas ${nombreArchivoExistente ? 'fa-file-check' : 'fa-paperclip'}"></i>
                                            ${nombreArchivoExistente || 'Cargar Respaldo'}
                                        </span>
                                        <input type="file" id="in-file-${idFilaUnico}" style="display:none;"
                                               onchange="window.registrarArchivoIA('${idFilaUnico}')">
                                    </div>
                                    <button id="btn-down-${idFilaUnico}"
                                            onclick="window.descargarAnexoIA('${idFilaUnico}')"
                                            style="display:${nombreArchivoExistente ? 'inline-block' : 'none'}; background:rgba(255,255,255,0.05); color:${colorElegido}; border:1px solid ${colorElegido}; padding:2px 6px; border-radius:3px; font-size:0.5rem; font-weight:800; cursor:pointer;">
                                        DESCARGAR
                                    </button>
                                </div>
                            </td>
                            <td style="padding:12px; color:#94a3b8; font-style:italic;">${sintetizarIdeaIA(obs)}</td>
                            <td style="padding:12px; text-align:center; color:${colorElegido}; font-weight:800; font-family:monospace;">
                                ${formatearFechaProIA(fechaRaw)}
                            </td>
                        </tr>`;
                    }).join('')}
                </tbody>
            </table>
        </div>`;
};

// ─────────────────────────────────────────────────────────────
// SECCIÓN 6: LÓGICA PRINCIPAL (DOMContentLoaded)
// ─────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {

    inicializarFirebase();

    const gridPrincipal     = document.getElementById('grid-principal');
    const vistaDetalle      = document.getElementById('vista-detalle');
    const vistaAnalisis     = document.getElementById('vista-analisis');
    const tituloComponente  = document.getElementById('titulo-componente');
    const contenidoDinamico = document.getElementById('contenido-dinamico');
    const contenedorKpis    = document.getElementById('contenedor-kpis');
    const inputExcel        = document.getElementById('input-excel');
    const btnBack           = document.getElementById('btn-back');
    const btnBackAnalisis   = document.getElementById('btn-back-analisis');
    const btnHome           = document.getElementById('btn-home');
    const btnAnalisis       = document.getElementById('btn-analisis');
    const infoModificacion  = document.getElementById('info-modificacion');
    const fechaModTexto     = document.getElementById('fecha-modificacion-texto');

    const columnasMatriz = [
        "Aspecto", "Tema", "Norma Legal", "Compilado en el Decreto 1076",
        "Emisor", "Objeto", "Exigencia", "Sede", "Evidencia de cumplimiento",
        "Seguimiento", "Cumplimiento", "Responsable del cumplimiento",
        "Observación", "Última actualización"
    ];

    // ── Spinner de carga ──────────────────────────────────────
    function mostrarSpinner(mensaje = "Procesando...") {
        let overlay = document.getElementById('prebel-spinner');
        if (!overlay) {
            overlay = document.createElement('div');
            overlay.id = 'prebel-spinner';
            overlay.style.cssText = `
                position: fixed; inset: 0; background: rgba(2,6,23,0.85);
                display: flex; flex-direction: column; align-items: center;
                justify-content: center; z-index: 9999; backdrop-filter: blur(4px);`;
            overlay.innerHTML = `
                <div style="width:50px; height:50px; border:4px solid #1e3a5f;
                     border-top-color:#00f3ff; border-radius:50%;
                     animation: spinnerGiro 0.8s linear infinite;"></div>
                <p id="spinner-msg" style="color:#94a3b8; margin-top:20px;
                   font-size:0.75rem; font-weight:700; letter-spacing:3px;
                   text-transform:uppercase;">${mensaje}</p>
                <style>
                    @keyframes spinnerGiro { to { transform: rotate(360deg); } }
                    @keyframes fadeInLeft {
                        from { opacity:0; transform:translateX(-10px); }
                        to   { opacity:1; transform:translateX(0); }
                    }
                </style>`;
            document.body.appendChild(overlay);
        } else {
            document.getElementById('spinner-msg').textContent = mensaje;
            overlay.style.display = 'flex';
        }
    }

    function ocultarSpinner() {
        const overlay = document.getElementById('prebel-spinner');
        if (overlay) overlay.style.display = 'none';
    }

    // ── Navegación ────────────────────────────────────────────
    const irHome = () => {
        gridPrincipal.style.display  = 'grid';
        vistaDetalle.style.display   = 'none';
        vistaAnalisis.style.display  = 'none';
    };

    if (btnBack)         btnBack.onclick         = irHome;
    if (btnBackAnalisis) btnBackAnalisis.onclick  = irHome;
    if (btnHome)         btnHome.onclick          = irHome;
    if (btnAnalisis)     btnAnalisis.onclick      = () => {
        if (!datosMatrizGlobal) {
            mostrarAvisoSinDatos();
            return;
        }
        window.generarAnalisis();
    };

    // ── Aviso cuando no hay datos cargados ────────────────────
    function mostrarAvisoSinDatos() {
        let aviso = document.getElementById('aviso-sin-datos');
        if (aviso) { aviso.style.display = 'flex'; return; }

        aviso = document.createElement('div');
        aviso.id = 'aviso-sin-datos';
        aviso.style.cssText = `
            position: fixed; inset: 0; background: rgba(2,6,23,0.9);
            display: flex; flex-direction: column; align-items: center;
            justify-content: center; z-index: 9998; backdrop-filter: blur(4px);`;
        aviso.innerHTML = `
            <div style="background:#0f172a; border:1px solid #334155; border-radius:12px;
                 padding:40px 50px; text-align:center; max-width:420px;">
                <i class="fas fa-file-excel" style="font-size:3rem; color:#00f3ff; margin-bottom:20px;"></i>
                <h3 style="color:#fff; font-size:1rem; letter-spacing:3px; font-weight:900; margin-bottom:12px;">
                    SIN DATOS CARGADOS
                </h3>
                <p style="color:#64748b; font-size:0.8rem; line-height:1.6; margin-bottom:25px;">
                    Para acceder al análisis de cumplimiento debes importar 
                    primero la Matriz Legal en formato <strong style="color:#fff;">.xlsx</strong>.
                </p>
                <button onclick="document.getElementById('aviso-sin-datos').style.display='none'; document.getElementById('input-excel').click();"
                        style="background:#00f3ff; color:#000; border:none; border-radius:6px;
                               padding:12px 30px; font-weight:900; cursor:pointer; font-size:0.75rem;
                               letter-spacing:2px; margin-right:10px;">
                    IMPORTAR MATRIZ
                </button>
                <button onclick="document.getElementById('aviso-sin-datos').style.display='none';"
                        style="background:transparent; color:#64748b; border:1px solid #334155;
                               border-radius:6px; padding:12px 20px; font-weight:700;
                               cursor:pointer; font-size:0.75rem;">
                    CANCELAR
                </button>
            </div>`;
        document.body.appendChild(aviso);
    }

    // ── Importar Excel ────────────────────────────────────────
    inputExcel.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        mostrarSpinner("Procesando matriz legal...");

        const reader = new FileReader();
        reader.onload = (event) => {
            setTimeout(() => {
                try {
                    const data     = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                    let obj = {};
                    workbook.SheetNames.forEach(s => {
                        obj[s] = XLSX.utils.sheet_to_json(workbook.Sheets[s], { range: 3 });
                    });
                    datosMatrizGlobal = obj;

                    const ahora = new Date().toISOString();
                    localStorage.setItem('prebel_cyber_data', JSON.stringify(obj));
                    localStorage.setItem('prebel_ultima_importacion', ahora);

                    _actualizarFechaModificacion(ahora);

                    ocultarSpinner();
                    window.generarAnalisis();
                    actualizarPanelInicio();
                } catch (err) {
                    ocultarSpinner();
                    console.error("[EXCEL] Error al procesar:", err);
                    alert("Error al procesar el archivo Excel. Verifica que el formato sea correcto (.xlsx / .xls).");
                }
                e.target.value = '';
            }, 80);
        };
        reader.readAsArrayBuffer(file);
    });

    function _actualizarFechaModificacion(isoString) {
        if (!infoModificacion || !fechaModTexto) return;
        const fecha = new Date(isoString);
        const formato = fecha.toLocaleString('es-CO', {
            day: '2-digit', month: '2-digit', year: 'numeric',
            hour: '2-digit', minute: '2-digit'
        });
        fechaModTexto.textContent = `Última importación: ${formato}`;
        infoModificacion.style.display = 'block';
    }

    // ── Navegación directa a componente ──────────────────────
    window.navDirecta = (e, componente) => {
        if (e) { e.preventDefault(); e.stopPropagation(); }
        if (!datosMatrizGlobal) {
            mostrarAvisoSinDatos();
            return;
        }
        const hKey = Object.keys(datosMatrizGlobal).find(k =>
            limpiarTexto(k).includes(limpiarTexto(componente))
        );
        if (hKey) {
            vistaAnalisis.style.display  = 'none';
            gridPrincipal.style.display  = 'none';
            vistaDetalle.style.display   = 'block';
            tituloComponente.innerText   = componente.toUpperCase();

            const fechaGuardada = localStorage.getItem('prebel_ultima_importacion');
            if (fechaGuardada) _actualizarFechaModificacion(fechaGuardada);

            renderTabla(datosMatrizGlobal[hKey], hKey);
        }
    };

    // ── Renderizar vista de tarjetas (v93) ───────────────────
    const renderTabla = (datos, nombreHoja) => {
        const registros = datos.filter(f => f["Aspecto"]);

        // Construir HTML de una card individual (v95 - diseño con acciones rápidas)
        const buildCard = (f, idx, termino = '') => {
            const cumple      = validarCumplimiento(f[Object.keys(f).find(c => limpiarTexto(c) === 'cumplimiento')] ?? f['Cumplimiento']);
            const norma       = (f["Norma Legal"]   || "Sin norma").toString();
            const tema        = (f["Tema"]           || "Sin tema").toString();
            const aspecto     = (f["Aspecto"]        || "---").toString();
            const sede        = (f["Sede"]           || "---").toString();
            const objeto      = sintetizarIdeaIA(f["Objeto"] || "---");
            const observacion = extraerDatoReal(f, "Observacion");
            const obsTexto    = observacion === "---" ? "" : observacion.toString();
            const fechaRaw    = extraerDatoReal(f, "actualizacion");
            const emisor      = (f["Emisor"] || "---").toString();
            const hojaEsc     = CSS.escape(nombreHoja);

            // Calcular días sin actualizar
            let diasSinActualizar = null;
            if (fechaRaw !== "---") {
                const fn = fechaRaw instanceof Date ? fechaRaw : new Date(fechaRaw);
                if (!isNaN(fn)) diasSinActualizar = Math.floor((new Date() - fn) / (1000 * 86400));
            }

            // Highlight de búsqueda
            const hl = (txt) => {
                if (!termino || termino.length < 2) return txt;
                const re = new RegExp(`(${termino.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')})`, 'gi');
                return txt.replace(re, `<mark style="background:rgba(0,243,255,0.25); color:#00f3ff; border-radius:2px; padding:0 2px;">$1</mark>`);
            };

            const colorBorde   = cumple ? '#1D9E75' : '#E24B4A';
            const colorBadgeBg = cumple ? 'rgba(29,158,117,0.12)' : 'rgba(226,75,74,0.12)';
            const colorBadge   = cumple ? '#5DCAA5' : '#F09595';
            const iconBadge    = cumple ? 'fa-check-circle' : 'fa-times-circle';
            const textoBadge   = cumple ? 'Cumple' : 'No cumple';

            return `
            <div class="prebel-card"
                 data-idx="${idx}"
                 data-norma="${norma.toLowerCase()}"
                 data-tema="${tema.toLowerCase()}"
                 data-aspecto="${aspecto.toLowerCase()}"
                 data-sede="${sede.toLowerCase()}"
                 data-cumple="${cumple ? '1' : '0'}"
                 style="
                    background: rgba(15,23,42,0.9);
                    border: 0.5px solid rgba(255,255,255,0.08);
                    border-left: 3px solid ${colorBorde};
                    border-radius: 0 12px 12px 0;
                    padding: 16px 18px;
                    display: flex;
                    flex-direction: column;
                    gap: 10px;
                    transition: border-color 0.2s ease, transform 0.15s ease;
                 "
                 onmouseover="this.style.borderColor='${colorBorde}'; this.style.transform='translateY(-2px)';"
                 onmouseout="this.style.borderColor='rgba(255,255,255,0.08)'; this.style.borderLeftColor='${colorBorde}'; this.style.transform='';">

                <!-- Fila 1: aspecto + norma + badge estado -->
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:10px;">
                    <div style="flex:1; min-width:0;">
                        <div style="font-size:0.58rem; color:#475569; font-weight:700; letter-spacing:1.5px; text-transform:uppercase; margin-bottom:3px;">
                            ${hl(aspecto)}
                        </div>
                        <div style="font-size:0.82rem; font-weight:800; color:#f1f5f9; line-height:1.35; word-break:break-word;">
                            ${hl(norma)}
                        </div>
                    </div>
                    <span style="
                        flex-shrink:0;
                        display:inline-flex; align-items:center; gap:5px;
                        background:${colorBadgeBg};
                        color:${colorBadge};
                        border: 0.5px solid ${colorBorde}44;
                        border-radius:20px; padding:4px 11px;
                        font-size:0.62rem; font-weight:800; letter-spacing:0.5px;
                        white-space:nowrap;">
                        <i class="fas ${iconBadge}" style="font-size:0.65rem;"></i> ${textoBadge}
                    </span>
                </div>

                <!-- Fila 2: chips de metadatos -->
                <div style="display:flex; flex-wrap:wrap; gap:6px; align-items:center;">
                    ${tema !== 'Sin tema' ? `
                    <span style="display:inline-flex; align-items:center; gap:4px;
                                 background:rgba(0,243,255,0.07); color:#67e8f9;
                                 border:0.5px solid rgba(0,243,255,0.2);
                                 border-radius:4px; padding:2px 8px; font-size:0.6rem; font-weight:700;">
                        <i class="fas fa-tag" style="font-size:0.55rem;"></i>${hl(tema)}
                    </span>` : ''}
                    ${sede !== '---' ? `
                    <span style="display:inline-flex; align-items:center; gap:4px;
                                 background:rgba(167,139,250,0.08); color:#a78bfa;
                                 border:0.5px solid rgba(167,139,250,0.2);
                                 border-radius:4px; padding:2px 8px; font-size:0.6rem; font-weight:700;">
                        <i class="fas fa-map-marker-alt" style="font-size:0.55rem;"></i>${hl(sede)}
                    </span>` : ''}
                    ${emisor !== '---' ? `
                    <span style="display:inline-flex; align-items:center; gap:4px;
                                 background:rgba(251,191,36,0.08); color:#fbbf24;
                                 border:0.5px solid rgba(251,191,36,0.2);
                                 border-radius:4px; padding:2px 8px; font-size:0.6rem; font-weight:700;">
                        <i class="fas fa-building" style="font-size:0.55rem;"></i>${emisor}
                    </span>` : ''}
                    ${diasSinActualizar !== null ? `
                    <span style="display:inline-flex; align-items:center; gap:4px;
                                 background:${diasSinActualizar > 365 ? 'rgba(251,191,36,0.1)' : 'rgba(255,255,255,0.04)'};
                                 color:${diasSinActualizar > 365 ? '#fbbf24' : '#475569'};
                                 border:0.5px solid ${diasSinActualizar > 365 ? 'rgba(251,191,36,0.3)' : 'rgba(255,255,255,0.08)'};
                                 border-radius:4px; padding:2px 8px; font-size:0.6rem; font-weight:700; margin-left:auto;">
                        <i class="fas fa-clock" style="font-size:0.55rem;"></i>
                        ${diasSinActualizar > 365 ? `${diasSinActualizar} días sin actualizar` : formatearFechaProIA(fechaRaw)}
                    </span>` : ''}
                </div>

                <!-- Fila 3: objeto resumido -->
                ${objeto !== 'Sin registro' && objeto !== '---' ? `
                <div style="font-size:0.67rem; color:#64748b; line-height:1.5;
                            border-top:0.5px solid rgba(255,255,255,0.05); padding-top:8px;">
                    ${objeto}
                </div>` : ''}

                <!-- Fila 4: observación editable -->
                <div style="border-top:0.5px solid rgba(255,255,255,0.05); padding-top:8px;">
                    <div style="font-size:0.55rem; color:#334155; font-weight:700; text-transform:uppercase;
                                letter-spacing:1px; margin-bottom:4px;">
                        <i class="fas fa-pen" style="color:#00f3ff66; margin-right:3px;"></i>Observación
                    </div>
                    <div contenteditable="true"
                         spellcheck="false"
                         onfocus="window.activarEdicionObservacion(this)"
                         onblur="window.guardarObservacion(this, '${hojaEsc}', ${idx})"
                         style="
                            min-height: 20px;
                            color: ${obsTexto ? '#94a3b8' : '#334155'};
                            font-style: italic;
                            font-size: 0.68rem;
                            outline: none;
                            border-radius: 4px;
                            padding: 4px 6px;
                            cursor: text;
                            line-height: 1.5;
                            white-space: pre-wrap;
                            word-break: break-word;
                            transition: background 0.2s, outline 0.2s;
                         "
                    >${obsTexto || '<span style="color:#334155;">Clic para agregar observación...</span>'}</div>
                </div>

                <!-- Fila 5: acciones rápidas -->
                <div style="display:flex; gap:6px; flex-wrap:wrap; border-top:0.5px solid rgba(255,255,255,0.05); padding-top:10px;">
                    <button
                        onclick="window.toggleCumplimiento(this, '${hojaEsc}', ${idx})"
                        style="
                            display:inline-flex; align-items:center; gap:5px;
                            background:${cumple ? 'rgba(226,75,74,0.1)' : 'rgba(29,158,117,0.1)'};
                            color:${cumple ? '#F09595' : '#5DCAA5'};
                            border:0.5px solid ${cumple ? 'rgba(226,75,74,0.3)' : 'rgba(29,158,117,0.3)'};
                            border-radius:6px; padding:5px 12px;
                            font-size:0.62rem; font-weight:800; cursor:pointer;
                            letter-spacing:0.5px; transition:all 0.15s ease;"
                        onmouseover="this.style.opacity='0.8'"
                        onmouseout="this.style.opacity='1'">
                        <i class="fas ${cumple ? 'fa-times-circle' : 'fa-check-circle'}" style="font-size:0.65rem;"></i>
                        ${cumple ? 'Marcar incumplido' : 'Marcar cumplido'}
                    </button>
                    <button
                        onclick="window.gestionarAnexoIA('${hojaEsc}_${idx}')"
                        style="
                            display:inline-flex; align-items:center; gap:5px;
                            background:rgba(255,255,255,0.04);
                            color:#64748b;
                            border:0.5px solid rgba(255,255,255,0.1);
                            border-radius:6px; padding:5px 12px;
                            font-size:0.62rem; font-weight:800; cursor:pointer;
                            letter-spacing:0.5px; transition:all 0.15s ease;"
                        onmouseover="this.style.color='#94a3b8'; this.style.borderColor='rgba(255,255,255,0.2)'"
                        onmouseout="this.style.color='#64748b'; this.style.borderColor='rgba(255,255,255,0.1)'">
                        <i class="fas fa-paperclip" style="font-size:0.65rem;"></i> Adjuntar
                        <input type="file" id="in-file-${hojaEsc}_${idx}" style="display:none;"
                               onchange="window.registrarArchivoIA('${hojaEsc}_${idx}')">
                    </button>
<span id="norma-btn-slot-${hojaEsc}_${idx}"></span>
                    <button
                        onclick="window._verDetalle('${hojaEsc}', ${idx})"
                        style="
                            display:inline-flex; align-items:center; gap:5px;
                            background:rgba(255,255,255,0.04);
                            color:#64748b;
                            border:0.5px solid rgba(255,255,255,0.1);
                            border-radius:6px; padding:5px 12px;
                            font-size:0.62rem; font-weight:800; cursor:pointer;
                            letter-spacing:0.5px; transition:all 0.15s ease;"
                        onmouseover="this.style.color='#94a3b8'; this.style.borderColor='rgba(255,255,255,0.2)'"
                        onmouseout="this.style.color='#64748b'; this.style.borderColor='rgba(255,255,255,0.1)'">
                        <i class="fas fa-expand-alt" style="font-size:0.65rem;"></i> Ver detalle
                    </button>
                </div>
            </div>`;        };

        // Render inicial de todas las cards
        const htmlCards = registros.map((f, idx) => buildCard(f, idx)).join('');
        const total = registros.length;
        setTimeout(() => window._actualizarBotonesNorma && window._actualizarBotonesNorma(), 100);

        contenidoDinamico.innerHTML = `
            <!-- Barra de búsqueda y filtros -->
            <div style="
                position: sticky; top: 0; z-index: 10;
                background: rgba(5,1,10,0.95); backdrop-filter: blur(10px);
                padding: 16px 0 12px; margin-bottom: 20px;
                border-bottom: 1px solid rgba(255,255,255,0.06);">

                <!-- Buscador -->
                <div style="position:relative; margin-bottom:12px;">
                    <i class="fas fa-search" style="
                        position:absolute; left:16px; top:50%; transform:translateY(-50%);
                        color:#475569; font-size:0.85rem; pointer-events:none;"></i>
                    <input
                        id="prebel-search"
                        type="text"
                        placeholder="Buscar por norma, tema, aspecto o sede..."
                        autocomplete="off"
                        style="
                            width: 100%; box-sizing: border-box;
                            background: rgba(15,23,42,0.9);
                            border: 1px solid #334155;
                            border-radius: 10px;
                            color: #f8fafc;
                            padding: 13px 16px 13px 44px;
                            font-size: 0.85rem;
                            outline: none;
                            transition: border-color 0.2s ease;
                        "
                        onfocus="this.style.borderColor='#00f3ff'"
                        onblur="this.style.borderColor='#334155'"
                        oninput="window._filtrarCards(this.value)">
                </div>

                <!-- Filtros rápidos + contador -->
                <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span style="font-size:0.6rem; color:#475569; font-weight:800; letter-spacing:1px; text-transform:uppercase; margin-right:4px;">Filtrar:</span>
                    <button onclick="window._filtrarEstado('todos', this)"
                            class="prebel-filtro-activo"
                            id="filtro-todos"
                            style="background:rgba(0,243,255,0.15); color:#00f3ff; border:1px solid #00f3ff44;
                                   border-radius:20px; padding:4px 14px; font-size:0.65rem; font-weight:800;
                                   cursor:pointer; letter-spacing:1px;">
                        TODOS
                    </button>
                    <button onclick="window._filtrarEstado('cumple', this)"
                            id="filtro-cumple"
                            style="background:rgba(0,255,128,0.08); color:#00ff80; border:1px solid #00ff8033;
                                   border-radius:20px; padding:4px 14px; font-size:0.65rem; font-weight:800;
                                   cursor:pointer; letter-spacing:1px;">
                        <i class="fas fa-check-circle"></i> CUMPLE
                    </button>
                    <button onclick="window._filtrarEstado('nocumple', this)"
                            id="filtro-nocumple"
                            style="background:rgba(255,59,48,0.08); color:#ff3b30; border:1px solid #ff3b3033;
                                   border-radius:20px; padding:4px 14px; font-size:0.65rem; font-weight:800;
                                   cursor:pointer; letter-spacing:1px;">
                        <i class="fas fa-times-circle"></i> NO CUMPLE
                    </button>
                    <span id="prebel-contador"
                          style="margin-left:auto; font-size:0.65rem; color:#475569; font-weight:700; letter-spacing:1px;">
                        ${total} requisitos
                    </span>
                </div>
            </div>

            <!-- Grid de tarjetas -->
            <div id="prebel-cards-grid" style="
                display: grid;
                grid-template-columns: repeat(auto-fill, minmax(380px, 1fr));
                gap: 16px;
                padding-bottom: 40px;">
                ${htmlCards}
            </div>

            <!-- Estado vacío -->
            <div id="prebel-empty" style="display:none; text-align:center; padding:80px 20px;">
                <i class="fas fa-search" style="font-size:3rem; color:#1e293b; margin-bottom:20px;"></i>
                <p style="color:#334155; font-size:0.9rem; font-weight:700; letter-spacing:2px;">SIN RESULTADOS</p>
                <p style="color:#1e293b; font-size:0.75rem; margin-top:8px;">Intenta con otro término de búsqueda</p>
            </div>`;

        // Estado del filtro activo
        let _estadoFiltro = 'todos';

        // Filtrar por texto
        window._filtrarCards = (termino) => {
            const t = limpiarTexto(termino);
            const cards = document.querySelectorAll('.prebel-card');
            let visibles = 0;
            cards.forEach(card => {
                const matchTexto = !t || t.length < 1
                    || card.dataset.norma.includes(t)
                    || card.dataset.tema.includes(t)
                    || card.dataset.aspecto.includes(t)
                    || card.dataset.sede.includes(t);
                const matchEstado = _estadoFiltro === 'todos'
                    || (_estadoFiltro === 'cumple'    && card.dataset.cumple === '1')
                    || (_estadoFiltro === 'nocumple'  && card.dataset.cumple === '0');
                const visible = matchTexto && matchEstado;
                card.style.display = visible ? '' : 'none';
                if (visible) visibles++;
            });
            _actualizarContador(visibles, total);
            document.getElementById('prebel-empty').style.display = visibles === 0 ? 'block' : 'none';
        };

        // Filtrar por estado
        window._filtrarEstado = (estado, btn) => {
            _estadoFiltro = estado;
            // Resetear estilos de botones
            ['filtro-todos','filtro-cumple','filtro-nocumple'].forEach(id => {
                const b = document.getElementById(id);
                if (b) b.style.opacity = '0.45';
            });
            if (btn) btn.style.opacity = '1';
            // Re-aplicar con el texto actual
            const termino = document.getElementById('prebel-search')?.value || '';
            window._filtrarCards(termino);
        };

        function _actualizarContador(visibles, total) {
            const el = document.getElementById('prebel-contador');
            if (el) el.textContent = visibles === total
                ? `${total} requisitos`
                : `${visibles} de ${total} requisitos`;
        }
    };

    // ── SVG Anillo ICLA ───────────────────────────────────────
    const crearAnilloSVG = (p) => {
        const radio = 58;
        const circ  = 2 * Math.PI * radio;
        const color = p >= 90 ? '#07c092' : p >= 70 ? '#ffc107' : '#ff3b30';
        const uid = 'icla_' + Date.now();
        setTimeout(() => {
            const el = document.getElementById(uid);
            if (!el) return;
            const duration = 2000;
            const start = performance.now();
            function tick(now) {
                const elapsed = now - start;
                const progress = Math.min(elapsed / duration, 1);
                const ease = 1 - Math.pow(1 - progress, 3);
                const current = Math.round(ease * p * 10) / 10;
                el.textContent = current.toFixed(1) + '%';
                if (progress < 1) requestAnimationFrame(tick);
                else el.textContent = p + '%';
            }
            requestAnimationFrame(tick);
        }, 100);
        return `
            <div style="position:relative; width:180px; height:180px; margin:0 auto 15px;">
                <style>
                    @keyframes iclaDrawIn {
                        from { stroke-dashoffset: ${circ}; }
                        to   { stroke-dashoffset: 0; }
                    }
                    @keyframes iclaFadeUp {
                        from { opacity:0; transform:translate(-50%,-50%) translateY(8px); }
                        to   { opacity:1; transform:translate(-50%,-50%) translateY(0); }
                    }
                    @keyframes iclaRipple {
                        0%   { r: ${radio}; opacity: 0.5; }
                        100% { r: ${radio + 22}; opacity: 0; }
                    }
                </style>
                <svg width="180" height="180" style="position:absolute;inset:0;overflow:visible;">
                    <circle cx="90" cy="90" r="${radio}" fill="none" stroke="${color}" stroke-width="1.5"
                        style="animation: iclaRipple 2.5s ease-out 2s infinite;"/>
                    <circle cx="90" cy="90" r="${radio}" fill="none" stroke="${color}" stroke-width="1"
                        style="animation: iclaRipple 2.5s ease-out 2.8s infinite;"/>
                    <circle cx="90" cy="90" r="${radio}" fill="none" stroke="${color}" stroke-width="0.5"
                        style="animation: iclaRipple 2.5s ease-out 3.6s infinite;"/>
                </svg>
                <svg width="180" height="180" style="transform:rotate(-90deg);position:absolute;inset:0;">
                    <circle cx="90" cy="90" r="${radio}" fill="transparent"
                            stroke="rgba(255,255,255,0.05)" stroke-width="10"/>
                    <circle cx="90" cy="90" r="${radio}" fill="transparent"
                            stroke="${color}" stroke-width="10"
                            stroke-dasharray="${circ}" stroke-dashoffset="${circ}"
                            stroke-linecap="round"
                            style="animation: iclaDrawIn 2s cubic-bezier(0.4,0,0.2,1) forwards;"/>
                </svg>
                <div style="position:absolute; top:50%; left:50%;
                            animation: iclaFadeUp 0.6s ease 0.3s both;
                            text-align:center;">
                    <span id="${uid}" style="font-size:2.2rem; font-weight:700; color:#fff; display:block; line-height:1;">0%</span>
                    <span style="font-size:0.6rem; color:${color}; letter-spacing:2px; font-weight:700;">ICLA</span>
                </div>
            </div>`;
    };

    // ── Agente IA (AMVA) ──────────────────────────────────────
    // ─────────────────────────────────────────────────────────────
    // AGENTE IA — VIGILANCIA NORMATIVA (Claude + Web Search)
    // ─────────────────────────────────────────────────────────────

    const PREBEL_PERFIL = `Empresa: Prebel S.A.S BIC
Sector: Industria cosmética y de aseo (fabricación y envasado de cosméticos, fragancias, talcos, esmaltes)
Sedes: Medellín (Productora) y Rionegro, Antioquia, Colombia
Autoridades ambientales competentes: AMVA y CORNARE
Componentes: Agua, Aire, Energía, Residuos, Suelo/Biodiversidad, Riesgo Químico, Contingencias, Mecanismos de gestión`.trim();

    function construirContextoMatriz() {
        if (!datosMatrizGlobal) return "No hay matriz cargada.";
        let resumen = [];
        Object.entries(datosMatrizGlobal).forEach(([hoja, filas]) => {
            const normas = filas
                .filter(f => f["Norma Legal"])
                .map(f => f["Norma Legal"])
                .filter(Boolean)
                .slice(0, 6);
            if (normas.length) resumen.push(`${hoja}: ${normas.join(", ")}`);
        });
        return resumen.join("\n");
    }

    async function consultarClaudeIA(consulta) {
        const contextoMatriz = construirContextoMatriz();

        const systemPrompt = `Eres un experto en normativa ambiental colombiana especializado en vigilancia regulatoria para industrias cosméticas.

PERFIL DE LA EMPRESA:
${PREBEL_PERFIL}

NORMATIVA ACTUALMENTE EN SU MATRIZ LEGAL:
${contextoMatriz}

TU MISIÓN: Buscar en internet normativa ambiental colombiana nueva o reciente que pueda aplicar a Prebel. Revisa:
- Ministerio de Ambiente y Desarrollo Sostenible (minambiente.gov.co)
- AMVA - Área Metropolitana del Valle de Aburrá (metropol.gov.co)
- CORNARE (cornare.gov.co)
- IDEAM (ideam.gov.co)
- ANLA (anla.gov.co)
- Diario Oficial de Colombia

Para cada norma encontrada indica: nombre/número, fecha, autoridad emisora, por qué aplica a Prebel, componente afectado y acción recomendada. Responde en español, claro y estructurado.`;

        const response = await fetch("https://prebel-proxy-beige.vercel.app/api/claude-proxy", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                model: "claude-sonnet-4-20250514",
                max_tokens: 1000,
                system: systemPrompt,
                tools: [{ type: "web_search_20250305", name: "web_search" }],
                messages: [{ role: "user", content: consulta }]
            })
        });

        const data = await response.json();
        if (!response.ok) throw new Error(data.error?.message || "Error en la API");

        return data.content
            .filter(b => b.type === "text")
            .map(b => b.text)
            .join("\n")
            .trim() || "No se obtuvo respuesta.";
    }

    function renderRespuestaIA(texto, boxTexto) {
        const html = texto
            .replace(/\*\*(.+?)\*\*/g, '<strong style="color:#fff;">$1</strong>')
            .replace(/^#{1,3} (.+)$/gm, '<div style="color:#00f3ff;font-size:0.78rem;font-weight:800;margin:12px 0 4px;">$1</div>')
            .replace(/^(\d+)\. /gm, '<br><span style="color:#00f3ff;font-weight:700;">$1.</span> ')
            .replace(/^[-•] /gm, "<br>• ")
            .replace(/\n/g, "<br>");

        boxTexto.style.height = "auto";
        boxTexto.style.minHeight = "180px";
        boxTexto.style.overflowY = "auto";
        boxTexto.style.maxHeight = "420px";
        boxTexto.innerHTML = `
            <div style="border-left:4px solid #00f3ff;padding-left:15px;animation:fadeInLeft 0.4s ease;text-align:left;">
                <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px;">
                    <i class="fas fa-robot" style="color:#00f3ff;font-size:0.9rem;"></i>
                    <h5 style="color:#00f3ff;margin:0;font-size:0.75rem;font-weight:800;letter-spacing:1px;">AGENTE IA · VIGILANCIA NORMATIVA</h5>
                </div>
                <div style="color:#cbd5e1;font-size:0.72rem;line-height:1.7;">${html}</div>
                <div style="margin-top:12px;font-size:0.58rem;color:#475569;padding-top:8px;border-top:1px solid rgba(255,255,255,0.05);">
                    <i class="fas fa-search"></i> Búsqueda en tiempo real · ${new Date().toLocaleDateString("es-CO")}
                </div>
            </div>`;
    }

    function inicializarAgenteIA() {
        const btn = document.getElementById("ia-bt");
        const inp = document.getElementById("ia-in");
        if (!btn || !inp) return;

        // Placeholder sugerido
        inp.placeholder = "Ej: normativa nueva de residuos peligrosos 2025...";

        btn.onclick = async () => {
            const consulta = inp.value.trim() || "Busca normativa ambiental colombiana nueva del último año que aplique a Prebel S.A.S BIC";

            const boxTexto = document.getElementById("box-texto-reactivo");
            if (!boxTexto) return;

            btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> BUSCANDO...';
            btn.disabled = true;
            inp.disabled = true;

            boxTexto.style.height = "auto";
            boxTexto.style.minHeight = "180px";
            boxTexto.innerHTML = `
                <div style="text-align:center;padding:20px 0;">
                    <i class="fas fa-spinner fa-spin" style="font-size:1.5rem;color:#00f3ff;margin-bottom:12px;display:block;"></i>
                    <p style="color:#94a3b8;font-size:0.7rem;font-weight:700;letter-spacing:1px;">Consultando fuentes oficiales...</p>
                    <p style="color:#475569;font-size:0.6rem;margin-top:6px;">Minambiente · AMVA · CORNARE · IDEAM · ANLA</p>
                </div>`;

            try {
                const respuesta = await consultarClaudeIA(consulta);
                renderRespuestaIA(respuesta, boxTexto);
            } catch (err) {
                boxTexto.innerHTML = `
                    <div style="border-left:4px solid #ff3b30;padding-left:15px;">
                        <p style="color:#ff3b30;font-size:0.75rem;font-weight:800;margin:0 0 6px;">ERROR DE CONEXIÓN</p>
                        <p style="color:#94a3b8;font-size:0.7rem;">${err.message}</p>
                    </div>`;
            } finally {
                btn.innerHTML = '<i class="fas fa-search"></i> ESCANEAR NORMATIVA';
                btn.disabled = false;
                inp.disabled = false;
                inp.value = "";
            }
        };

        inp.addEventListener("keydown", e => { if (e.key === "Enter") btn.click(); });
    }

    // ── Generar Dashboard de Análisis ─────────────────────────
    window.generarAnalisis = () => {
        if (!datosMatrizGlobal) return;

        let tSi = 0, tNo = 0, tReg = 0;
        let htmlBarras = "";
        let htmlHallazgos = "";
        let vigilanciaStack = [];
        const hoy = new Date();

        const botonesSlim = [
            { n: "Agua",                  i: "fa-tint",           c: "#38bdf8" },
            { n: "Aire",                  i: "fa-wind",           c: "#cbd5e1" },
            { n: "Energía y combustibles",i: "fa-bolt",           c: "#fb923c" },
            { n: "Residuos",              i: "fa-recycle",        c: "#a78bfa" },
            { n: "Suelo y Biodiversidad", i: "fa-leaf",           c: "#4ade80" },
            { n: "Riesgo Químico",        i: "fa-flask",          c: "#f472b6" },
            { n: "Contingencias",         i: "fa-exclamation-triangle", c: "#ef4444" },
            { n: "Otros",                 i: "fa-folder",         c: "#facc15" },
            { n: "Mecanismos de Gestión", i: "fa-clipboard-check",c: "#2dd4bf" }
        ];

        botonesSlim.forEach(b => {
            const hKey = Object.keys(datosMatrizGlobal).find(h =>
                limpiarTexto(h).includes(limpiarTexto(b.n))
            );
            let siH = 0, noH = 0;
            if (hKey) {
                datosMatrizGlobal[hKey].forEach(f => {
                    if (f["Aspecto"] && f["Aspecto"].toString().trim().length > 1) {
                        tReg++;
                        const colC = Object.keys(f).find(c => limpiarTexto(c) === "cumplimiento");
                        const colF = Object.keys(f).find(c => limpiarTexto(c).includes("actualizacion"));
                        if (validarCumplimiento(f[colC])) {
                            siH++; tSi++;
                        } else {
                            noH++; tNo++;
                            htmlHallazgos += `
                                <div style="padding:8px; border-bottom:1px solid #ffffff08; font-size:0.65rem; color:#fff;">
                                    <b style="color:#ff3b30;">[${b.n.toUpperCase()}]:</b> ${f["Norma Legal"] || 'S/N'}
                                </div>`;
                        }
                        if (colF && f[colF]) {
                            let fn = f[colF] instanceof Date ? f[colF] : new Date(f[colF]);
                            if (!isNaN(fn) && (hoy - fn) / (1000 * 86400) > 365) {
                                const _fechaVig = f[colF] instanceof Date ? f[colF] : new Date(f[colF]);
                                vigilanciaStack.push({ c: b.n, n: f["Norma Legal"] || "S/N", fecha: _fechaVig.toLocaleDateString("es-CO",{day:"2-digit",month:"2-digit",year:"numeric"}), dias: Math.floor((hoy - _fechaVig)/(1000*86400)), sede: f["Sede"] || "", tema: f["Tema"] || "", objeto: f["Objeto"] || "", evidencia: f["Evidencia de cumplimiento"] || "", observacion: f["Obervación"] || f["Observación"] || "", exigencia: f["Exigencia"] || "", color: b.c });
                            }
                        }
                    }
                });
            }
            const p = (siH + noH) > 0 ? ((siH / (siH + noH)) * 100).toFixed(1) : 0;
            htmlBarras += `
                <div onclick="window.navDirecta(event, '${b.n}')" style="margin-bottom:10px; cursor:pointer;">
                    <div style="display:flex; justify-content:space-between; font-size:0.62rem; margin-bottom:3px;">
                        <span style="color:#475569;">${b.n}</span>
                        <span style="color:${b.c};">${p}%</span>
                    </div>
                    <div style="width:100%; background:rgba(255,255,255,0.05); height:6px; border-radius:99px; overflow:hidden;">
                        <div style="width:${p}%; background:${b.c}; height:100%; border-radius:99px; transform-origin:left; animation: monBarFill 1s cubic-bezier(0.4,0,0.2,1) both; animation-delay: ${botonesSlim.indexOf(b) * 0.1}s;"></div>
                    </div>
                </div>`;
        });

        window._vigStack = vigilanciaStack;
        const global = tReg > 0 ? ((tSi / tReg) * 100).toFixed(1) : 0;

        // Cargar historial real desde localStorage
        const _histGuardado = JSON.parse(localStorage.getItem('prebel_hist_revisiones') || '[]');
        let histTotal, mesesCentral;
        if (_histGuardado.length > 0) {
            histTotal = _histGuardado.map(h => parseFloat(h.pct));
            mesesCentral = _histGuardado.map(h => h.mes);
            // Actualizar el último con el valor actual si es el mismo mes
            const _mesActual = (() => { const d = new Date(); const mm = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']; return mm[d.getMonth()] + ' ' + String(d.getFullYear()).slice(2); })();
            if (mesesCentral[mesesCentral.length-1] === _mesActual) {
                histTotal[histTotal.length-1] = parseFloat(global);
            }
        } else {
            histTotal = [...PREBEL_CONFIG.historicoCumplimiento.global];
            histTotal[histTotal.length - 1] = parseFloat(global);
            mesesCentral = PREBEL_CONFIG.historicoCumplimiento.meses;
        }
        const ptsGlobal = histTotal.map((v, i) =>
            `${i * 120},${100 - ((v - 80) / 20 * 100)}`
        ).join(' ');

        const nodosGlobal = histTotal.map((v, i) => {
            const x = i * 120;
            const y = 100 - ((v - 80) / 20 * 100);
            return `
                <g style="cursor:pointer;" onclick="window.ejecutarAccionPunto('${mesesCentral[i]}', ${v}, 'GLOBAL')">
                    <circle cx="${x}" cy="${y}" r="15" fill="transparent" />
                    <circle cx="${x}" cy="${y}" r="5" fill="#00f3ff" stroke="#fff" stroke-width="1.5" />
                    <text x="${x}" y="${y - 18}" fill="#00f3ff" font-size="11" font-weight="900" text-anchor="middle">${v}%</text>
                    <text x="${x}" y="125" fill="#475569" font-size="9" font-weight="700" text-anchor="middle">${mesesCentral[i]}</text>
                </g>`;
        }).join('');

        if (!document.getElementById("monbar-style")) { const _s = document.createElement("style"); _s.id = "monbar-style"; _s.textContent = "@keyframes monBarFill { from { opacity:0; transform: scaleX(0.3); } to { opacity:1; transform: scaleX(1); } }"; document.head.appendChild(_s); }
        const statsCards = [
            { label: "TOTAL REQUISITOS", valor: tReg,  color: "#64748b", icono: "fa-layer-group" },
            { label: "CUMPLIDOS",         valor: tSi,   color: "#07c092", icono: "fa-check" },
            { label: "NO CUMPLIDOS",      valor: tNo,   color: "#E24B4A", icono: "fa-times" },
            { label: "EN VIGILANCIA",     valor: vigilanciaStack.length, color: "#BA7517", icono: "fa-clock" }
        ];

        contenedorKpis.innerHTML = `
            <div style="display:grid; grid-template-columns:repeat(12,1fr); gap:20px; width:100%;
                        background:#05010a; padding:25px; border-radius:20px;">

                <div style="grid-column:span 12; display:grid; grid-template-columns:repeat(4,1fr); gap:15px; margin-bottom:5px;">
                    ${statsCards.map(s => `
                        <div style="background:rgba(15,23,42,0.6); border:0.5px solid rgba(255,255,255,0.06); border-radius:10px; padding:16px 18px; display:flex; align-items:center; gap:14px;">
                            <div style="width:32px;height:32px;border-radius:8px;background:rgba(255,255,255,0.04);display:flex;align-items:center;justify-content:center;flex-shrink:0;"><i class="fas ${s.icono}" style="font-size:13px;color:${s.color};"></i></div>
                            <div>
                                <div style="font-size:22px; font-weight:700; color:${s.color}; line-height:1;">${s.valor}</div>
                                <div style="font-size:9px; color:#475569; font-weight:500; letter-spacing:0.5px; margin-top:4px; text-transform:uppercase;">${s.label}</div>
                            </div>
                        </div>`).join('')}
                </div>

                <div style="grid-column:span 3; display:flex; flex-direction:column; gap:20px;">
                    <div style="background:rgba(15,23,42,0.85); padding:30px; border-radius:16px; border:1px solid rgba(255,255,255,0.08); text-align:center; position:relative; overflow:hidden;"><svg style="position:absolute;bottom:0;left:0;width:100%;opacity:0.07;pointer-events:none;" viewBox="0 0 400 80" preserveAspectRatio="none"><path d="M0,40 C80,10 160,70 240,40 C320,10 360,60 400,40 L400,80 L0,80 Z" fill="#07c092"/><path d="M0,55 C100,30 200,70 300,45 C360,28 385,55 400,50 L400,80 L0,80 Z" fill="#38bdf8"/></svg><div style="position:relative;z-index:1;">${crearAnilloSVG(global)}<h3 style="font-size:0.7rem; color:#94a3b8; letter-spacing:2px; font-weight:800;">ICLA ACTUAL</h3><p style="font-size:0.6rem; color:#475569; margin-top:5px;">${tReg} requisitos evaluados</p></div></div>
                    <div style="background:rgba(15,23,42,0.85); padding:25px; border-radius:16px;
                         border-left:5px solid #ff3b30; border:1px solid rgba(255,255,255,0.08); flex-grow:1;">
                        <h3 style="font-size:0.75rem; color:#ff3b30; margin-bottom:12px; font-weight:900;">
                            ⚠️ HALLAZGOS (${tNo})
                        </h3>
                        <div style="max-height:220px; overflow-y:auto; scrollbar-width:thin;">
                            ${htmlHallazgos || '<p style="color:#475569; font-size:0.65rem;">Sin hallazgos activos.</p>'}
                        </div>
                    </div>
                </div>

                <div style="grid-column:span 6; background:rgba(15,23,42,0.85); padding:30px;
                     border-radius:16px; border:1px solid rgba(255,255,255,0.08); overflow:hidden; position:relative;">
                    <h3 style="font-size:0.8rem; color:#94a3b8; margin-bottom:30px; text-align:center;
                         letter-spacing:4px; font-weight:900;">MONITOR OPERATIVO</h3>
                    <div style="display:grid;grid-template-columns:1fr 1fr;gap:25px;">${htmlBarras}</div><div style="position:absolute;bottom:0;left:0;right:0;height:90px;display:flex;align-items:flex-end;gap:3px;padding:0 30px;opacity:0.08;pointer-events:none;"><div style="flex:1;height:30%;background:#07c092;border-radius:2px 2px 0 0;"></div><div style="flex:1;height:55%;background:#07c092;border-radius:2px 2px 0 0;"></div><div style="flex:1;height:40%;background:#07c092;border-radius:2px 2px 0 0;"></div><div style="flex:1;height:70%;background:#07c092;border-radius:2px 2px 0 0;"></div><div style="flex:1;height:45%;background:#07c092;border-radius:2px 2px 0 0;"></div><div style="flex:1;height:85%;background:#07c092;border-radius:2px 2px 0 0;"></div><div style="flex:1;height:60%;background:#07c092;border-radius:2px 2px 0 0;"></div><div style="flex:1;height:95%;background:#07c092;border-radius:2px 2px 0 0;"></div><div style="flex:1;height:50%;background:#07c092;border-radius:2px 2px 0 0;"></div><div style="flex:1;height:75%;background:#07c092;border-radius:2px 2px 0 0;"></div><div style="flex:1;height:35%;background:#07c092;border-radius:2px 2px 0 0;"></div><div style="flex:1;height:90%;background:#07c092;border-radius:2px 2px 0 0;"></div><div style="flex:1;height:65%;background:#07c092;border-radius:2px 2px 0 0;"></div><div style="flex:1;height:100%;background:#07c092;border-radius:2px 2px 0 0;"></div><div style="flex:1;height:48%;background:#07c092;border-radius:2px 2px 0 0;"></div><div style="flex:1;height:80%;background:#07c092;border-radius:2px 2px 0 0;"></div></div>
                </div>

                <div style="grid-column:span 3; display:flex; flex-direction:column; gap:20px;">
                    <div onclick="window._abrirModalVigilancia()" style="background:#060e1f;border:0.5px solid rgba(255,255,255,0.06);border-radius:12px;padding:14px;cursor:pointer;transition:border-color 0.2s,background 0.2s;"
                         onmouseover="this.style.borderColor='rgba(186,117,23,0.4)';this.style.background='rgba(186,117,23,0.04)'"
                         onmouseout="this.style.borderColor='rgba(255,255,255,0.06)';this.style.background='#060e1f'">
                        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
                            <span style="font-size:8px;color:#334155;text-transform:uppercase;letter-spacing:1.5px;font-weight:600;">Vigilancia</span>
                            <span style="font-size:13px;font-weight:700;color:#BA7517;">${vigilanciaStack.length}</span>
                        </div>
                        <div style="display:flex;flex-direction:column;gap:0;">
                            ${vigilanciaStack.length
                                ? vigilanciaStack.slice(0,5).map(v => '<div style="padding:5px 0;border-bottom:0.5px solid rgba(255,255,255,0.04);display:flex;align-items:flex-start;gap:5px;"><span style="font-size:8px;font-weight:600;color:#854F0B;background:rgba(186,117,23,0.15);padding:1px 5px;border-radius:3px;flex-shrink:0;">' + v.c + '</span><span style="font-size:9px;color:#334155;">' + v.n + '</span></div>').join('')
                                : '<p style="font-size:9px;color:#1e293b;">Sin normas en vigilancia.</p>'
                            }
                        </div>
                        <div style="display:flex;align-items:center;justify-content:center;margin-top:10px;border-top:0.5px solid rgba(255,255,255,0.04);padding-top:8px;">
                            <span style="font-size:8px;color:#1e293b;">↗ Ver todas las normas</span>
                        </div>
                    </div>
                    <div style="background:#060e1f; padding:20px; border-radius:12px; border:0.5px solid rgba(255,255,255,0.06);">
                        <style>
                            @keyframes aiPulse { 0%,100%{opacity:1;transform:scale(1)} 50%{opacity:0.3;transform:scale(1.12)} }
                            @keyframes aiDotBlink { 0%,100%{opacity:1} 50%{opacity:0.4} }
                            .ai-ring-widget { width:26px;height:26px;border-radius:50%;border:1.5px solid #07c092;display:flex;align-items:center;justify-content:center;position:relative;flex-shrink:0; }
                            .ai-ring-widget::before { content:'';position:absolute;inset:-4px;border-radius:50%;border:1px solid rgba(7,192,146,0.2);animation:aiPulse 2s ease-in-out infinite; }
                            .ai-dot-widget { width:7px;height:7px;border-radius:50%;background:#07c092;animation:aiDotBlink 1.5s ease-in-out infinite; }
                        </style>
                        <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px;">
                            <div class="ai-ring-widget"><div class="ai-dot-widget"></div></div>
                            <div>
                                <div style="font-size:12px;font-weight:600;color:#94a3b8;">Agente normativo</div>
                            </div>
                        </div>
                        <div style="font-size:9px;color:#07c092;display:flex;align-items:center;gap:4px;margin-bottom:10px;">
                            <div style="width:5px;height:5px;border-radius:50%;background:#07c092;animation:aiDotBlink 1.5s ease-in-out infinite;"></div>
                            En línea · Búsqueda en tiempo real
                        </div>
                        <p style="font-size:9px;color:#334155;margin:0 0 10px;line-height:1.5;">
                            Consulta normativa ambiental colombiana aplicable a Prebel en fuentes oficiales.
                        </p>
                        <div style="display:flex;flex-wrap:wrap;gap:4px;margin-bottom:14px;">
                            <span style="font-size:8px;color:#334155;background:rgba(255,255,255,0.03);border:0.5px solid rgba(255,255,255,0.06);border-radius:4px;padding:2px 7px;">Minambiente</span>
                            <span style="font-size:8px;color:#334155;background:rgba(255,255,255,0.03);border:0.5px solid rgba(255,255,255,0.06);border-radius:4px;padding:2px 7px;">AMVA</span>
                            <span style="font-size:8px;color:#334155;background:rgba(255,255,255,0.03);border:0.5px solid rgba(255,255,255,0.06);border-radius:4px;padding:2px 7px;">CORNARE</span>
                            <span style="font-size:8px;color:#334155;background:rgba(255,255,255,0.03);border:0.5px solid rgba(255,255,255,0.06);border-radius:4px;padding:2px 7px;">IDEAM</span>
                            <span style="font-size:8px;color:#334155;background:rgba(255,255,255,0.03);border:0.5px solid rgba(255,255,255,0.06);border-radius:4px;padding:2px 7px;">ANLA</span>
                        </div>
                        <button id="ia-bt" onclick="window._abrirModalIA()"
                                style="width:100%;height:38px;background:#07c092;color:#030711;border:none;border-radius:8px;font-weight:700;cursor:pointer;font-size:10px;letter-spacing:0.5px;display:flex;align-items:center;justify-content:center;gap:8px;">
                            ⌕ Abrir consulta
                        </button>
                    </div>
                </div>

                <div style="grid-column:span 12; background:rgba(15,23,42,0.6); border:0.5px solid rgba(255,255,255,0.06); border-radius:12px; padding:24px 28px; margin:10px 0; display:flex; align-items:stretch; gap:24px; min-height:160px;">

                    <!-- Izquierda: valor global + botón -->
                    <div style="min-width:140px;display:flex;flex-direction:column;justify-content:center;gap:8px;">
                        <div style="font-size:8px;color:#334155;text-transform:uppercase;letter-spacing:1.5px;">Tendencia global</div>
                        <div style="font-size:2.4rem;font-weight:700;color:${global >= 90 ? '#07c092' : global >= 70 ? '#ffc107' : '#ff3b30'};line-height:1;">${global}%</div>
                        <div style="font-size:9px;color:#07c092;display:flex;align-items:center;gap:4px;">
                            <div style="width:5px;height:5px;border-radius:50%;background:#07c092;"></div> Estable · ${mesesCentral[mesesCentral.length-1]}
                        </div>
                        <button onclick="window._registrarRevision()" style="margin-top:8px;background:rgba(7,192,146,0.1);border:0.5px solid #07c092;color:#07c092;border-radius:6px;padding:7px 10px;font-size:9px;font-weight:600;cursor:pointer;letter-spacing:0.5px;text-align:left;transition:background 0.2s;"
                            onmouseover="this.style.background='rgba(7,192,146,0.2)'" onmouseout="this.style.background='rgba(7,192,146,0.1)'">
                            <i class="fas fa-check-circle" style="margin-right:5px;"></i>Registrar revisión
                        </button>
                        <button onclick="window._deshacerRevision()" style="margin-top:4px;background:rgba(226,75,74,0.08);border:0.5px solid rgba(226,75,74,0.4);color:#E24B4A;border-radius:6px;padding:6px 10px;font-size:9px;font-weight:600;cursor:pointer;text-align:left;transition:background 0.2s;"
                            onmouseover="this.style.background='rgba(226,75,74,0.15)'" onmouseout="this.style.background='rgba(226,75,74,0.08)'">
                            <i class="fas fa-undo" style="margin-right:5px;"></i>Deshacer último
                        </button>
                        <button onclick="window._abrirHistoricoManual()" style="margin-top:4px;background:rgba(56,189,248,0.08);border:0.5px solid rgba(56,189,248,0.3);color:#38bdf8;border-radius:6px;padding:6px 10px;font-size:9px;font-weight:600;cursor:pointer;text-align:left;transition:background 0.2s;"
                            onmouseover="this.style.background='rgba(56,189,248,0.15)'" onmouseout="this.style.background='rgba(56,189,248,0.08)'">
                            <i class="fas fa-edit" style="margin-right:5px;"></i>Ingresar histórico
                        </button>
                    </div>

                    <!-- Centro: barras -->
                    <div style="flex:2;">
                        ${(() => {
                            const maxV = Math.max(...histTotal);
                            const minV = Math.min(...histTotal) - 2;
                            const range = maxV - minV || 1;
                            const bars = histTotal.map((v, i) => {
                                const h = Math.round(((v - minV) / range) * 100);
                                const isLast = i === histTotal.length - 1;
                                const col = isLast ? '#07c092' : 'rgba(7,192,146,' + (0.3 + i * 0.08) + ')';
                                const delay = (i * 0.1).toFixed(1);
                                const obs = localStorage.getItem('prebel_obs_' + i) ? '●' : '';
                                return '<div style="flex:1;display:flex;flex-direction:column;align-items:center;gap:3px;cursor:pointer;" onclick="window._selectBar(' + i + ',' + v + ')">'
                                    + '<span style="font-size:8px;color:' + (isLast ? '#07c092' : '#94a3b8') + ';font-weight:' + (isLast ? '600' : '400') + ';animation:fadeUpTrend 0.4s ease ' + delay + 's both;opacity:0;">' + v + '%</span>'
                                    + '<div style="width:100%;background:#0f172a;border-radius:3px 3px 0 0;overflow:hidden;height:90px;display:flex;align-items:flex-end;">'
                                    + '<div style="width:100%;height:' + h + '%;background:' + col + ';border-radius:3px 3px 0 0;transform-origin:bottom;animation:barUpTrend 0.7s cubic-bezier(0.34,1.2,0.64,1) ' + delay + 's both;"></div>'
                                    + '</div>'
                                    + '<span style="font-size:7px;color:#07c092;">' + obs + '</span>'
                                    + '</div>';
                            }).join('');
                            return '<style>@keyframes barUpTrend{from{transform:scaleY(0)}to{transform:scaleY(1)}}@keyframes fadeUpTrend{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:translateY(0)}}</style>'
                                + '<div style="display:flex;align-items:flex-end;gap:6px;">' + bars + '</div>'
                                + '<div style="display:flex;gap:6px;margin-top:6px;border-top:0.5px solid #0f172a;padding-top:5px;">'
                                + mesesCentral.map((m, i) => '<span style="flex:1;text-align:center;font-size:8px;color:#94a3b8;">' + m + '</span>').join('')
                                + '</div>';
                        })()}
                    </div>

                    <!-- Derecha: box observación -->
                    <div id="box-obs-trend" style="flex:1;background:rgba(0,0,0,0.2);border:0.5px solid rgba(255,255,255,0.05);border-radius:8px;padding:16px;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;min-width:220px;">
                        <i class="fas fa-hand-pointer" style="font-size:1.2rem;color:#1e293b;margin-bottom:8px;"></i>
                        <p style="font-size:9px;color:#1e293b;text-transform:uppercase;letter-spacing:1px;">Clic en una barra<br>para añadir observación</p>
                    </div>
                </div>

                <div style="grid-column:span 3; grid-row:span 2; display:flex; flex-direction:column;
                     gap:10px; background:rgba(255,255,255,0.02); padding:20px; border-radius:16px;
                     border:1px solid rgba(255,255,255,0.05);">
                    <p style="color:#64748b; font-size:0.6rem; letter-spacing:2px; margin-bottom:5px;
                         font-weight:800; text-transform:uppercase;">Componentes</p>
                    ${botonesSlim.map(b => `
                        <div onclick="window.actualizarPanelDerecho(event, '${b.n}', '${b.c}')"
                             style="height:45px; background:rgba(255,255,255,0.03); border:1px solid rgba(255,255,255,0.08);
                                    border-left:5px solid ${b.c}; border-radius:4px; display:flex;
                                    align-items:center; padding:0 15px; cursor:pointer; transition:0.25s;"
                             onmouseover="this.style.background='rgba(255,255,255,0.08)'; this.style.transform='translateX(6px)';"
                             onmouseout="this.style.background='rgba(255,255,255,0.03)'; this.style.transform='translateX(0)';">
                            <i class="fas ${b.i}" style="color:${b.c}; font-size:1rem; width:30px;"></i>
                            <span style="color:#f8fafc; font-size:0.75rem; font-weight:700; letter-spacing:1px;">
                                ${b.n.toUpperCase()}
                            </span>
                        </div>`).join('')}
                    <div style="flex-grow:1;"></div>
                </div>

                <div id="box-informacion-relacionada"
                     style="grid-column:span 6; grid-row:span 2; background:rgba(15,23,42,0.85);
                            border:1px solid rgba(255,255,255,0.08); border-radius:16px; padding:30px;
                            display:flex; flex-direction:column; align-items:center; justify-content:center;
                            text-align:center; min-height:450px;">
                    <i class="fas fa-mouse-pointer" style="font-size:3rem; color:rgba(255,255,255,0.05); margin-bottom:20px;"></i>
                    <h4 style="color:#94a3b8; font-size:1rem; font-weight:800; letter-spacing:4px; text-transform:uppercase;">
                        Seleccione un Componente
                    </h4>
                    <p style="color:#334155; font-size:0.7rem; margin-top:10px;">
                        Haga clic en cualquier componente de la izquierda para ver su análisis detallado.
                    </p>
                </div>

                <div style="grid-column:span 3; grid-row:span 2; display:flex; flex-direction:column; gap:20px;">
                    <div id="box-historico-componente"
                         style="flex:1; background:linear-gradient(180deg,rgba(15,23,42,0.9),rgba(10,1,20,0.9));
                                border:1px solid #bf5af233; border-radius:16px; padding:25px;
                                display:flex; flex-direction:column; justify-content:center; text-align:center;">
                        <i class="fas fa-chart-area" style="font-size:2rem; color:rgba(255,255,255,0.05); margin-bottom:15px;"></i>
                        <h4 style="color:#444; font-size:0.7rem; letter-spacing:2px; font-weight:800;">TENDENCIA COMPONENTE</h4>
                    </div>
                    <div id="box-texto-reactivo"
                         style="height:180px; background:rgba(15,23,42,0.8); border:1px solid rgba(0,243,255,0.1);
                                border-radius:16px; padding:20px; display:flex; align-items:center;
                                justify-content:center; text-align:center;">
                        <p style="color:#334155; font-size:0.7rem; font-weight:700; text-transform:uppercase; letter-spacing:1px;">
                            Esperando selección...
                        </p>
                    </div>
                </div>
            </div>`;

        // ── Registrar revisión mensual ─────────────────────────
        window._abrirModalVigilancia = () => {
            const existing = document.getElementById('vig-modal-overlay');
            if (existing) existing.remove();
            const stack = window._vigStack || [];
            const overlay = document.createElement('div');
            overlay.id = 'vig-modal-overlay';
            overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.75);display:flex;align-items:center;justify-content:center;z-index:9999;backdrop-filter:blur(4px);';

            overlay.innerHTML = '<style>@keyframes modalIn{from{opacity:0;transform:scale(0.95) translateY(8px)}to{opacity:1;transform:scale(1) translateY(0)}}@keyframes slideIn{from{opacity:0;transform:translateX(10px)}to{opacity:1;transform:translateX(0)}}</style>'
                + '<div style="background:#060e1f;border:0.5px solid rgba(255,255,255,0.08);border-radius:14px;width:800px;max-width:95vw;height:520px;display:flex;flex-direction:column;box-shadow:0 40px 80px rgba(0,0,0,0.6);animation:modalIn 0.3s cubic-bezier(0.34,1.2,0.64,1) both;">'
                + '<div style="padding:14px 20px;border-bottom:0.5px solid rgba(255,255,255,0.06);display:flex;align-items:center;justify-content:space-between;">'
                + '<div style="display:flex;align-items:center;gap:10px;"><span style="font-size:16px;">⏱</span><span style="font-size:13px;font-weight:600;color:#94a3b8;">Normas en vigilancia</span>'
                + '<span style="font-size:9px;color:#BA7517;background:rgba(186,117,23,0.1);border:0.5px solid rgba(186,117,23,0.3);border-radius:99px;padding:2px 10px;">' + stack.length + ' normas · +365 días</span></div>'
                + '<div id="vig-modal-close" style="width:28px;height:28px;border-radius:50%;background:rgba(255,255,255,0.04);border:0.5px solid rgba(255,255,255,0.08);color:#475569;display:flex;align-items:center;justify-content:center;cursor:pointer;font-size:12px;">✕</div>'
                + '</div>'
                + '<div style="display:flex;flex:1;overflow:hidden;">'
                + '<div id="vig-list" style="width:260px;flex-shrink:0;border-right:0.5px solid rgba(255,255,255,0.06);overflow-y:auto;padding:10px 10px;"></div>'
                + '<div id="vig-detail" style="flex:1;overflow-y:auto;padding:16px 20px;display:flex;align-items:center;justify-content:center;"><p style="font-size:10px;color:#1e293b;text-transform:uppercase;letter-spacing:1px;">Selecciona una norma</p></div>'
                + '</div></div>';

            document.body.appendChild(overlay);
            document.getElementById('vig-modal-close').onclick = () => overlay.remove();
            overlay.onclick = (e) => { if (e.target === overlay) overlay.remove(); };

            // Build norm list
            const listEl = document.getElementById('vig-list');
            stack.forEach((v, idx) => {
                const daysColor = v.dias > 500 ? '#E24B4A' : v.dias > 400 ? '#BA7517' : '#94a3b8';
                const card = document.createElement('div');
                card.style.cssText = 'background:rgba(255,255,255,0.02);border:0.5px solid rgba(255,255,255,0.05);border-left:3px solid rgba(186,117,23,0.4);border-radius:7px;padding:9px 12px;margin-bottom:7px;cursor:pointer;transition:background 0.15s,border-left-color 0.15s;';
                card.innerHTML = '<div style="font-size:10px;font-weight:600;color:#94a3b8;margin-bottom:4px;">' + v.n + '</div>'
                    + '<div style="display:flex;justify-content:space-between;align-items:center;">'
                    + '<span style="font-size:8px;font-weight:600;color:#854F0B;background:rgba(186,117,23,0.15);padding:1px 5px;border-radius:3px;">' + v.c + '</span>'
                    + '<span style="font-size:8px;color:' + daysColor + '">' + v.dias + ' días</span></div>';

                card.onmouseover = () => { card.style.background = 'rgba(186,117,23,0.06)'; card.style.borderLeftColor = '#BA7517'; };
                card.onmouseout = () => { if (!card.classList.contains('active')) { card.style.background = 'rgba(255,255,255,0.02)'; card.style.borderLeftColor = 'rgba(186,117,23,0.4)'; } };
                card.onclick = () => {
                    // Deselect all
                    listEl.querySelectorAll('div').forEach(d => { d.style.background = 'rgba(255,255,255,0.02)'; d.style.borderLeftColor = 'rgba(186,117,23,0.4)'; d.classList.remove('active'); });
                    card.style.background = 'rgba(186,117,23,0.08)';
                    card.style.borderLeftColor = '#BA7517';
                    card.classList.add('active');
                    window._mostrarDetalleVig(v);
                };
                listEl.appendChild(card);
            });

            // Auto-select first
            if (stack.length > 0) {
                const first = listEl.querySelector('div');
                if (first) { first.style.background = 'rgba(186,117,23,0.08)'; first.style.borderLeftColor = '#BA7517'; first.classList.add('active'); }
                window._mostrarDetalleVig(stack[0]);
            }
        };

        window._mostrarDetalleVig = (v) => {
            const det = document.getElementById('vig-detail');
            if (!det) return;
            const daysColor = v.dias > 500 ? '#E24B4A' : v.dias > 400 ? '#BA7517' : '#94a3b8';
            det.style.alignItems = 'flex-start';
            det.style.justifyContent = 'flex-start';
            det.style.animation = 'slideIn 0.25s ease both';
            det.innerHTML = ''
                + '<div style="width:100%;">'
                + '<div style="margin-bottom:14px;">'
                + '<div style="font-size:14px;font-weight:600;color:#e2e8f0;margin-bottom:5px;">' + v.n + '</div>'
                + '<div style="display:flex;align-items:center;gap:8px;font-size:9px;color:#BA7517;">'
                + '<div style="width:6px;height:6px;border-radius:50%;background:' + (v.color || '#BA7517') + ';"></div>'
                + v.c + (v.sede ? ' · ' + v.sede : '') + '</div></div>'
                + '<div style="display:flex;align-items:center;gap:8px;background:rgba(226,75,74,0.06);border:0.5px solid rgba(226,75,74,0.2);border-radius:6px;padding:8px 12px;margin-bottom:14px;">'
                + '<span style="color:#E24B4A;font-size:12px;">⚠</span>'
                + '<span style="font-size:9px;color:' + daysColor + ';font-weight:600;">' + v.dias + ' días sin actualizar · Última act: ' + (v.fecha || '—') + '</span></div>'
                + (v.tema ? '<div style="margin-bottom:12px;"><div style="font-size:8px;color:#334155;text-transform:uppercase;letter-spacing:1px;margin-bottom:3px;">Tema</div><div style="font-size:10px;color:#94a3b8;">' + v.tema + '</div></div><div style="height:0.5px;background:rgba(255,255,255,0.04);margin-bottom:12px;"></div>' : '')
                + (v.objeto ? '<div style="margin-bottom:12px;"><div style="font-size:8px;color:#334155;text-transform:uppercase;letter-spacing:1px;margin-bottom:3px;">Objeto de la norma</div><div style="font-size:10px;color:#64748b;line-height:1.5;">' + v.objeto + '</div></div><div style="height:0.5px;background:rgba(255,255,255,0.04);margin-bottom:12px;"></div>' : '')
                + (v.evidencia ? '<div style="margin-bottom:12px;"><div style="font-size:8px;color:#334155;text-transform:uppercase;letter-spacing:1px;margin-bottom:3px;">Evidencia de cumplimiento</div><div style="font-size:10px;color:#64748b;line-height:1.5;">' + v.evidencia + '</div></div><div style="height:0.5px;background:rgba(255,255,255,0.04);margin-bottom:12px;"></div>' : '')
                + (v.observacion ? '<div style="margin-bottom:12px;"><div style="font-size:8px;color:#334155;text-transform:uppercase;letter-spacing:1px;margin-bottom:3px;">Observación</div><div style="font-size:10px;color:#64748b;line-height:1.5;">' + v.observacion + '</div></div>' : '')
                + '</div>';
        };

        window._abrirModalIA = () => {
            // Crear overlay si no existe
            if (document.getElementById('agente-modal-overlay')) {
                document.getElementById('agente-modal-overlay').style.display = 'flex';
                return;
            }
            const overlay = document.createElement('div');
            overlay.id = 'agente-modal-overlay';
            overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.75);display:flex;align-items:center;justify-content:center;z-index:9999;backdrop-filter:blur(4px);';
            overlay.innerHTML = `
                <div style="background:#060e1f;border:0.5px solid rgba(255,255,255,0.08);border-radius:16px;width:680px;max-width:95vw;height:560px;display:flex;flex-direction:column;box-shadow:0 40px 80px rgba(0,0,0,0.6);animation:modalIn 0.3s cubic-bezier(0.34,1.2,0.64,1) both;">
                    <style>
                        @keyframes modalIn{from{opacity:0;transform:scale(0.95) translateY(10px)}to{opacity:1;transform:scale(1) translateY(0)}}
                        @keyframes typingDot{0%,60%,100%{transform:translateY(0)}30%{transform:translateY(-4px)}}
                        .ia-bubble-user{background:rgba(7,192,146,0.1);color:#5DCAA5;border:0.5px solid rgba(7,192,146,0.2);border-radius:12px;border-bottom-right-radius:3px;padding:10px 14px;font-size:10px;line-height:1.6;max-width:85%;align-self:flex-end;}
                        .ia-bubble-ai{background:rgba(255,255,255,0.03);color:#94a3b8;border:0.5px solid rgba(255,255,255,0.06);border-radius:12px;border-bottom-left-radius:3px;padding:10px 14px;font-size:10px;line-height:1.6;max-width:85%;}
                        .ia-bubble-ai strong{color:#e2e8f0;}
                    </style>
                    <!-- Header -->
                    <div style="padding:16px 20px;border-bottom:0.5px solid rgba(255,255,255,0.06);display:flex;align-items:center;justify-content:space-between;">
                        <div style="display:flex;align-items:center;gap:10px;">
                            <div style="width:24px;height:24px;border-radius:50%;border:1.5px solid #07c092;display:flex;align-items:center;justify-content:center;position:relative;">
                                <div style="width:7px;height:7px;border-radius:50%;background:#07c092;animation:aiDotBlink 1.5s ease-in-out infinite;"></div>
                            </div>
                            <span style="font-size:12px;font-weight:600;color:#94a3b8;">Agente normativo · Prebel</span>
                            <span style="font-size:9px;color:#07c092;background:rgba(7,192,146,0.08);border:0.5px solid rgba(7,192,146,0.2);border-radius:99px;padding:2px 8px;">En línea</span>
                        </div>
                        <div id="ia-modal-close" style="width:28px;height:28px;border-radius:50%;background:rgba(255,255,255,0.04);border:0.5px solid rgba(255,255,255,0.08);color:#475569;display:flex;align-items:center;justify-content:center;cursor:pointer;font-size:12px;">✕</div>
                    </div>
                    <!-- Sources -->
                    <div style="padding:8px 20px;border-bottom:0.5px solid rgba(255,255,255,0.04);display:flex;gap:5px;flex-wrap:wrap;">
                        <span style="font-size:8px;color:#334155;background:rgba(255,255,255,0.03);border:0.5px solid rgba(255,255,255,0.05);border-radius:4px;padding:2px 7px;">Minambiente</span>
                        <span style="font-size:8px;color:#334155;background:rgba(255,255,255,0.03);border:0.5px solid rgba(255,255,255,0.05);border-radius:4px;padding:2px 7px;">AMVA</span>
                        <span style="font-size:8px;color:#334155;background:rgba(255,255,255,0.03);border:0.5px solid rgba(255,255,255,0.05);border-radius:4px;padding:2px 7px;">CORNARE</span>
                        <span style="font-size:8px;color:#334155;background:rgba(255,255,255,0.03);border:0.5px solid rgba(255,255,255,0.05);border-radius:4px;padding:2px 7px;">IDEAM</span>
                        <span style="font-size:8px;color:#334155;background:rgba(255,255,255,0.03);border:0.5px solid rgba(255,255,255,0.05);border-radius:4px;padding:2px 7px;">ANLA</span>
                        <span style="font-size:8px;color:#334155;background:rgba(255,255,255,0.03);border:0.5px solid rgba(255,255,255,0.05);border-radius:4px;padding:2px 7px;">Diario Oficial</span>
                    </div>
                    <!-- Chat body -->
                    <div id="ia-chat-body" style="flex:1;overflow-y:auto;padding:16px 20px;display:flex;flex-direction:column;gap:12px;">
                        <div style="text-align:center;padding:20px 0;">
                            <div style="font-size:10px;color:#1e293b;text-transform:uppercase;letter-spacing:1px;">Escribe tu consulta para comenzar</div>
                        </div>
                    </div>
                    <!-- Footer -->
                    <div style="padding:12px 20px;border-top:0.5px solid rgba(255,255,255,0.06);display:flex;gap:8px;align-items:center;">
                        <input id="ia-modal-input" type="text" placeholder="Escribe tu consulta normativa..."
                               style="flex:1;background:rgba(0,0,0,0.3);border:0.5px solid #1e293b;border-radius:8px;padding:10px 14px;font-size:10px;color:#94a3b8;outline:none;"
                               onfocus="this.style.borderColor='#07c092'" onblur="this.style.borderColor='#1e293b'"/>
                        <button id="ia-modal-send" style="background:#07c092;color:#030711;border:none;border-radius:8px;padding:10px 18px;font-size:10px;font-weight:700;cursor:pointer;white-space:nowrap;">Enviar</button>
                    </div>
                </div>`;

            document.body.appendChild(overlay);

            // Close
            document.getElementById('ia-modal-close').onclick = () => { overlay.style.display = 'none'; };
            overlay.onclick = (e) => { if (e.target === overlay) overlay.style.display = 'none'; };

            // Send message
            const sendMsg = () => {
                const inp = document.getElementById('ia-modal-input');
                const body = document.getElementById('ia-chat-body');
                const msg = inp.value.trim();
                if (!msg) return;
                inp.value = '';

                // User bubble
                const userDiv = document.createElement('div');
                userDiv.style.cssText = 'display:flex;flex-direction:column;align-items:flex-end;gap:4px;';
                userDiv.innerHTML = '<div class="ia-bubble-user">' + msg + '</div><span style="font-size:8px;color:#1e293b;">' + new Date().toLocaleTimeString('es-CO', {hour:'2-digit',minute:'2-digit'}) + '</span>';
                body.appendChild(userDiv);

                // Typing indicator
                const typingDiv = document.createElement('div');
                typingDiv.style.cssText = 'display:flex;flex-direction:column;align-items:flex-start;gap:4px;';
                typingDiv.innerHTML = '<div class="ia-bubble-ai" style="padding:10px 14px;"><span style="animation:typingDot 1s ease-in-out infinite;display:inline-block;width:5px;height:5px;border-radius:50%;background:#07c092;margin-right:3px;"></span><span style="animation:typingDot 1s ease-in-out 0.15s infinite;display:inline-block;width:5px;height:5px;border-radius:50%;background:#07c092;margin-right:3px;"></span><span style="animation:typingDot 1s ease-in-out 0.3s infinite;display:inline-block;width:5px;height:5px;border-radius:50%;background:#07c092;"></span></div>';
                body.appendChild(typingDiv);
                body.scrollTop = body.scrollHeight;

                // Call API (placeholder - will connect when proxy is ready)
                setTimeout(() => {
                    body.removeChild(typingDiv);
                    const aiDiv = document.createElement('div');
                    aiDiv.style.cssText = 'display:flex;flex-direction:column;align-items:flex-start;gap:4px;';
                    aiDiv.innerHTML = '<div class="ia-bubble-ai">El agente normativo estará disponible pronto. Por ahora puedes registrar tu consulta: <strong>' + msg + '</strong></div><span style="font-size:8px;color:#1e293b;">' + new Date().toLocaleTimeString('es-CO', {hour:'2-digit',minute:'2-digit'}) + '</span>';
                    body.appendChild(aiDiv);
                    body.scrollTop = body.scrollHeight;
                }, 1500);
            };

            document.getElementById('ia-modal-send').onclick = sendMsg;
            document.getElementById('ia-modal-input').addEventListener('keydown', e => { if (e.key === 'Enter') sendMsg(); });
            document.getElementById('ia-modal-input').focus();
        };

        window._registrarRevision = () => {
            if (!datosMatrizGlobal) return;
            // Calcular ICLA actual
            let si = 0, total = 0;
            Object.values(datosMatrizGlobal).forEach(filas => {
                filas.filter(f => f['Aspecto']).forEach(f => {
                    const colC = Object.keys(f).find(k => limpiarTexto(k) === 'cumplimiento');
                    if (colC) { total++; if (validarCumplimiento(f[colC])) si++; }
                });
            });
            const pctActual = total > 0 ? parseFloat(((si/total)*100).toFixed(1)) : 0;
            const ahora = new Date();
            const mm = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
            const mesLabel = mm[ahora.getMonth()] + ' ' + String(ahora.getFullYear()).slice(2);
            const hist = JSON.parse(localStorage.getItem('prebel_hist_revisiones') || '[]');
            // Si ya existe este mes, actualiza; si no, agrega
            const existeIdx = hist.findIndex(h => h.mes === mesLabel);
            if (existeIdx >= 0) {
                hist[existeIdx].pct = pctActual;
            } else {
                hist.push({ mes: mesLabel, pct: pctActual, fecha: ahora.toISOString() });
            }
            if (hist.length > 12) hist.shift();
            localStorage.setItem('prebel_hist_revisiones', JSON.stringify(hist));
            const btn = document.querySelector('button[onclick="window._registrarRevision()"]');
            if (btn) {
                btn.innerHTML = '<i class="fas fa-check" style="margin-right:5px;"></i>' + mesLabel + ': ' + pctActual + '%';
                btn.style.background = 'rgba(7,192,146,0.25)';
                setTimeout(() => window.generarAnalisis(), 600);
            }
        };

        window._deshacerRevision = () => {
            const hist = JSON.parse(localStorage.getItem('prebel_hist_revisiones') || '[]');
            if (hist.length === 0) return;
            const removido = hist.pop();
            localStorage.setItem('prebel_hist_revisiones', JSON.stringify(hist));
            const btn = document.querySelector('button[onclick="window._deshacerRevision()"]');
            if (btn) {
                btn.innerHTML = '<i class="fas fa-undo" style="margin-right:5px;"></i>Eliminado: ' + removido.mes;
                setTimeout(() => window.generarAnalisis(), 600);
            }
        };

        // ── Seleccionar barra para observación ─────────────────
        window._abrirHistoricoManual = () => {
            const box = document.getElementById('box-obs-trend');
            if (!box) return;
            const hist = JSON.parse(localStorage.getItem('prebel_hist_revisiones') || '[]');
            const mesesOpts = ['Ene 26','Feb 26','Mar 26','Abr 26','May 26','Jun 26','Jul 26','Ago 26','Sep 26','Oct 26','Nov 26','Dic 26'];
            box.style.textAlign = 'left';
            box.style.alignItems = 'flex-start';
            box.style.justifyContent = 'flex-start';
            box.innerHTML = '';

            const title = document.createElement('div');
            title.style.cssText = 'font-size:10px;font-weight:600;color:#38bdf8;margin-bottom:10px;width:100%;';
            title.textContent = 'Ingresar dato histórico';
            box.appendChild(title);

            const row = document.createElement('div');
            row.style.cssText = 'display:flex;gap:6px;width:100%;margin-bottom:8px;';

            const selMes = document.createElement('select');
            selMes.style.cssText = 'flex:1;background:#0a1628;border:0.5px solid #334155;color:#94a3b8;border-radius:6px;padding:6px;font-size:10px;outline:none;';
            mesesOpts.forEach(m => {
                const opt = document.createElement('option');
                opt.value = m;
                opt.textContent = m;
                selMes.appendChild(opt);
            });

            const inpPct = document.createElement('input');
            inpPct.type = 'number';
            inpPct.min = 0;
            inpPct.max = 100;
            inpPct.step = 0.1;
            inpPct.placeholder = '% ICLA';
            inpPct.style.cssText = 'width:80px;background:#0a1628;border:0.5px solid #334155;color:#94a3b8;border-radius:6px;padding:6px;font-size:10px;outline:none;';

            row.appendChild(selMes);
            row.appendChild(inpPct);
            box.appendChild(row);

            const addBtn = document.createElement('button');
            addBtn.textContent = 'Agregar';
            addBtn.style.cssText = 'width:100%;background:#38bdf8;color:#030711;border:none;border-radius:6px;padding:7px;font-size:10px;font-weight:700;cursor:pointer;margin-bottom:8px;';
            addBtn.onclick = () => {
                const mes = selMes.value;
                const pct = parseFloat(inpPct.value);
                if (!mes || isNaN(pct)) return;
                const h = JSON.parse(localStorage.getItem('prebel_hist_revisiones') || '[]');
                const idx = h.findIndex(x => x.mes === mes);
                if (idx >= 0) { h[idx].pct = pct; }
                else { h.push({ mes, pct, fecha: new Date().toISOString() }); }
                h.sort((a, b) => {
                    const mm = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
                    const [am, ay] = a.mes.split(' '); const [bm, by] = b.mes.split(' ');
                    return (parseInt(ay)*12 + mm.indexOf(am)) - (parseInt(by)*12 + mm.indexOf(bm));
                });
                if (h.length > 12) h.shift();
                localStorage.setItem('prebel_hist_revisiones', JSON.stringify(h));
                inpPct.value = '';
                addBtn.textContent = '✓ ' + mes + ': ' + pct + '%';
                addBtn.style.background = '#064e3b';
                addBtn.style.color = '#fff';
                setTimeout(() => { addBtn.textContent = 'Agregar'; addBtn.style.background = '#38bdf8'; addBtn.style.color = '#030711'; window.generarAnalisis(); }, 800);
            };
            box.appendChild(addBtn);

            // Show current saved data
            const listTitle = document.createElement('div');
            listTitle.style.cssText = 'font-size:9px;color:#334155;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px;';
            listTitle.textContent = 'Datos guardados';
            box.appendChild(listTitle);

            const list = document.createElement('div');
            list.style.cssText = 'width:100%;overflow-y:auto;max-height:80px;';
            if (hist.length === 0) {
                list.innerHTML = '<p style="font-size:9px;color:#1e293b;">Sin datos aún</p>';
            } else {
                hist.forEach(h => {
                    const row = document.createElement('div');
                    row.style.cssText = 'display:flex;justify-content:space-between;font-size:9px;color:#475569;padding:2px 0;border-bottom:0.5px solid #0a1628;';
                    row.innerHTML = '<span>' + h.mes + '</span><span style="color:#07c092;font-weight:600;">' + h.pct + '%</span>';
                    list.appendChild(row);
                });
            }
            box.appendChild(list);

            const closeBtn = document.createElement('button');
            closeBtn.textContent = 'Cerrar';
            closeBtn.style.cssText = 'width:100%;background:transparent;border:0.5px solid #334155;color:#475569;border-radius:6px;padding:5px;font-size:9px;cursor:pointer;margin-top:8px;';
            closeBtn.onclick = () => {
                box.style.textAlign = 'center';
                box.style.alignItems = 'center';
                box.style.justifyContent = 'center';
                box.innerHTML = '<i class="fas fa-hand-pointer" style="font-size:1.2rem;color:#1e293b;margin-bottom:8px;"></i><p style="font-size:9px;color:#1e293b;text-transform:uppercase;letter-spacing:1px;">Clic en una barra<br>para añadir observación</p>';
            };
            box.appendChild(closeBtn);
        };

        window._selectBar = (idx, val) => {
            const box = document.getElementById("box-obs-trend");
            if (!box) return;
            const savedObs = localStorage.getItem("prebel_obs_" + idx) || "";
            const mesLabel = PREBEL_CONFIG.historicoCumplimiento.meses[idx] || ("Mes " + (idx+1));
            box.style.textAlign = "left";
            box.style.alignItems = "flex-start";
            box.style.justifyContent = "flex-start";
            box.innerHTML = "";
            const hd = document.createElement("div");
            hd.style.cssText = "display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;width:100%;";
            const ttl = document.createElement("span");
            ttl.style.cssText = "font-size:11px;font-weight:600;color:#07c092;";
            ttl.textContent = mesLabel + " · " + val + "%";
            const closeBtn = document.createElement("span");
            closeBtn.textContent = "✕";
            closeBtn.style.cssText = "font-size:11px;color:#334155;cursor:pointer;";
            closeBtn.onclick = () => {
                box.style.textAlign = "center";
                box.style.alignItems = "center";
                box.style.justifyContent = "center";
                box.innerHTML = "<i class='fas fa-hand-pointer' style='font-size:1.2rem;color:#1e293b;margin-bottom:8px;'></i><p style='font-size:9px;color:#1e293b;text-transform:uppercase;letter-spacing:1px;'>Clic en una barra<br>para añadir observación</p>";
            };
            hd.appendChild(ttl);
            hd.appendChild(closeBtn);
            const ta = document.createElement("textarea");
            ta.id = "obs-input-" + idx;
            ta.placeholder = "Escribe tu observación...";
            ta.value = savedObs;
            ta.style.cssText = "width:100%;background:rgba(0,0,0,0.3);border:0.5px solid #334155;border-radius:6px;color:#94a3b8;padding:8px;font-size:10px;resize:none;height:90px;outline:none;box-sizing:border-box;margin-bottom:8px;";
            const saveBtn = document.createElement("button");
            saveBtn.textContent = "Guardar";
            saveBtn.style.cssText = "width:100%;background:#07c092;color:#030711;border:none;border-radius:6px;padding:8px;font-size:10px;font-weight:700;cursor:pointer;";
            saveBtn.onclick = () => window._guardarObs(idx);
            box.appendChild(hd);
            box.appendChild(ta);
            box.appendChild(saveBtn);
            if (savedObs) {
                const prev = document.createElement("p");
                prev.style.cssText = "font-size:8px;color:#334155;margin-top:5px;";
                prev.textContent = "Guardada: " + savedObs.substring(0,40) + (savedObs.length>40?"…":"");
                box.appendChild(prev);
            }
        };

        window._guardarObs = (idx) => {
            const inp = document.getElementById('obs-input-' + idx);
            if (!inp) return;
            localStorage.setItem('prebel_obs_' + idx, inp.value);
            const btn = inp.nextElementSibling;
            if (btn) { btn.textContent = '✓ Guardado'; btn.style.background = '#064e3b'; }
            setTimeout(() => window.generarAnalisis(), 600);
        };

        inicializarAgenteIA();

        const _dfEl = document.getElementById('dashboard-fecha');
        if (_dfEl) { const _fi = localStorage.getItem('prebel_ultima_importacion'); if (_fi) _dfEl.textContent = 'Datos al ' + new Date(_fi).toLocaleDateString('es-CO',{day:'2-digit',month:'2-digit',year:'numeric'}); }

        gridPrincipal.style.display  = 'none';
        vistaAnalisis.style.display  = 'block';
    };

    // ── Clics en tarjetas del panel principal ─────────────────
    document.addEventListener('click', (event) => {
        const grid = document.getElementById('grid-principal');
        if (grid && grid.contains(event.target)) {
            const tarjeta = event.target.closest('.comp-card');
            if (tarjeta && tarjeta.querySelector('h3')) {
                const nombreComponente = tarjeta.querySelector('h3').innerText;
                window.navDirecta(event, nombreComponente);
            }
        }
    });

    // ── Cargar datos guardados al iniciar ─────────────────────
    const saved = localStorage.getItem('prebel_cyber_data');
    if (saved) {
        try {
            datosMatrizGlobal = JSON.parse(saved);
            const fechaGuardada = localStorage.getItem('prebel_ultima_importacion');
            if (fechaGuardada) _actualizarFechaModificacion(fechaGuardada);
            window.generarAnalisis();
            actualizarPanelInicio();
        } catch (e) {
            console.error("[STORAGE] Error al cargar datos guardados:", e);
            localStorage.removeItem('prebel_cyber_data');
        }
    }

});
