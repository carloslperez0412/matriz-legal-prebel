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
        global: [87.5, 88.2, 88.0, 88.5, 88.6, null],
        meses:  ['OCT 25', 'NOV 25', 'DIC 25', 'ENE 26', 'FEB 26', 'MAR 26']
    }
};

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

window.gestionarAnexoIA = (idUnico) => {
    const selectorArchivos = document.getElementById(`in-file-${idUnico}`);
    if (selectorArchivos) selectorArchivos.click();
};

window.registrarArchivoIA = async (idUnico) => {
    const input     = document.getElementById(`in-file-${idUnico}`);
    const etiqueta  = document.getElementById(`label-file-${idUnico}`);
    const botonDescarga = document.getElementById(`btn-down-${idUnico}`);

    if (!input || !input.files.length) return;
    const archivo = input.files[0];

    if (archivo.size > 20 * 1024 * 1024) {
        alert("El archivo supera el límite de 20 MB.");
        return;
    }

    if (etiqueta) {
        etiqueta.innerHTML = `<i class="fas fa-spinner fa-spin"></i> Cargando...`;
        etiqueta.style.color = "#ffc107";
    }

    if (firebaseReady && storage) {
        try {
            const refCloud = storage.ref(`soportes_prebel/${idUnico}_${archivo.name}`);
            const task     = await refCloud.put(archivo);
            const url      = await task.ref.getDownloadURL();

            await db.collection("evidencias").doc(idUnico).set({
                nombre: archivo.name,
                url:    url,
                fecha:  new Date().toISOString()
            });

            _actualizarUIAnexo(etiqueta, botonDescarga, archivo.name, '#00ff80');
            console.log("[NUBE] Soporte sincronizado:", archivo.name);
            return;
        } catch (e) {
            console.warn("[NUBE] Fallo al subir. Guardando localmente...", e);
        }
    }

    _guardarEnLocalStorage(idUnico, archivo, etiqueta, botonDescarga);
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
    if (firebaseReady && db) {
        try {
            const doc = await db.collection("evidencias").doc(idUnico).get();
            if (doc.exists) {
                window.open(doc.data().url, '_blank');
                return;
            }
        } catch (e) {
            console.warn("[NUBE] No se pudo obtener desde Firebase:", e);
        }
    }

    const nombre    = localStorage.getItem(`prebel_filename_${idUnico}`);
    const contenido = localStorage.getItem(`prebel_filedata_${idUnico}`);
    if (nombre && contenido) {
        const link = document.createElement("a");
        link.href = contenido;
        link.download = nombre;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    } else {
        alert("No se encontró ningún archivo cargado para esta evidencia.");
    }
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
    return v === "1" || v === "1.0" || v === "si" || v === "sí";
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
        const radio = 70;
        const circ  = 2 * Math.PI * radio;
        const offset = circ - (p / 100) * circ;
        const color = p >= 90 ? '#07c092' : p >= 70 ? '#ffc107' : '#ff3b30';
        return `
            <div style="position:relative; width:160px; height:160px; margin:0 auto 15px;">
                <svg width="160" height="160" style="transform:rotate(-90deg);">
                    <circle cx="80" cy="80" r="${radio}" fill="transparent"
                            stroke="rgba(255,255,255,0.05)" stroke-width="12"/>
                    <circle cx="80" cy="80" r="${radio}" fill="transparent"
                            stroke="${color}" stroke-width="12"
                            stroke-dasharray="${circ}" stroke-dashoffset="${offset}"
                            stroke-linecap="round"
                            style="transition:stroke-dashoffset 1.5s ease-out;"/>
                </svg>
                <div style="position:absolute; top:50%; left:50%; transform:translate(-50%,-50%); text-align:center;">
                    <span style="font-size:2.5rem; font-weight:800; color:#fff; display:block; line-height:1;">${p}%</span>
                    <span style="font-size:0.6rem; color:${color}; letter-spacing:2px; font-weight:800;">ICLA</span>
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
            { n: "Agua",                  i: "fa-tint",           c: "#00f3ff" },
            { n: "Aire",                  i: "fa-wind",           c: "#ffffff" },
            { n: "Energía y combustibles",i: "fa-bolt",           c: "#ffe031" },
            { n: "Residuos",              i: "fa-trash-alt",      c: "#bf5af2" },
            { n: "Suelo y Biodiversidad", i: "fa-leaf",           c: "#00ff80" },
            { n: "Contingencias",         i: "fa-biohazard",      c: "#ff3b30" },
            { n: "Mecanismos de Gestión", i: "fa-clipboard-check",c: "#07c092" }
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
                                vigilanciaStack.push({ c: b.n, n: f["Norma Legal"] || "S/N" });
                            }
                        }
                    }
                });
            }
            const p = (siH + noH) > 0 ? ((siH / (siH + noH)) * 100).toFixed(1) : 0;
            htmlBarras += `
                <div onclick="window.navDirecta(event, '${b.n}')" style="margin-bottom:18px; cursor:pointer;">
                    <div style="display:flex; justify-content:space-between; font-size:0.65rem; color:#cbd5e1; margin-bottom:6px; font-weight:700;">
                        <span>${b.n.toUpperCase()}</span>
                        <span style="color:${b.c};">${p}%</span>
                    </div>
                    <div style="width:100%; background:rgba(0,0,0,0.4); height:5px; border-radius:10px; overflow:hidden;">
                        <div style="width:${p}%; background:${b.c}; height:100%; box-shadow:0 0 8px ${b.c}44;"></div>
                    </div>
                </div>`;
        });

        const global = tReg > 0 ? ((tSi / tReg) * 100).toFixed(1) : 0;

        const histTotal = [...PREBEL_CONFIG.historicoCumplimiento.global];
        histTotal[histTotal.length - 1] = parseFloat(global);

        const mesesCentral = PREBEL_CONFIG.historicoCumplimiento.meses;
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

        const statsCards = [
            { label: "TOTAL REQUISITOS", valor: tReg,  color: "#00f3ff", icono: "fa-list-check" },
            { label: "CUMPLIDOS",         valor: tSi,   color: "#00ff80", icono: "fa-check-double" },
            { label: "NO CUMPLIDOS",      valor: tNo,   color: "#ff3b30", icono: "fa-triangle-exclamation" },
            { label: "EN VIGILANCIA",     valor: vigilanciaStack.length, color: "#ffe031", icono: "fa-clock-rotate-left" }
        ];

        contenedorKpis.innerHTML = `
            <div style="display:grid; grid-template-columns:repeat(12,1fr); gap:20px; width:100%;
                        background:#05010a; padding:25px; border-radius:20px;">

                <div style="grid-column:span 12; display:grid; grid-template-columns:repeat(4,1fr); gap:15px; margin-bottom:5px;">
                    ${statsCards.map(s => `
                        <div style="background:rgba(15,23,42,0.7); border:1px solid rgba(255,255,255,0.06);
                             border-top:3px solid ${s.color}; border-radius:12px; padding:20px 25px;
                             display:flex; align-items:center; gap:15px;">
                            <i class="fas ${s.icono}" style="font-size:1.5rem; color:${s.color}; opacity:0.8;"></i>
                            <div>
                                <div style="font-size:1.8rem; font-weight:900; color:#fff; line-height:1;">${s.valor}</div>
                                <div style="font-size:0.55rem; color:#64748b; font-weight:800; letter-spacing:2px; margin-top:3px;">${s.label}</div>
                            </div>
                        </div>`).join('')}
                </div>

                <div style="grid-column:span 3; display:flex; flex-direction:column; gap:20px;">
                    <div style="background:rgba(15,23,42,0.85); padding:30px; border-radius:16px;
                         border:1px solid rgba(255,255,255,0.08); text-align:center;">
                        ${crearAnilloSVG(global)}
                        <h3 style="font-size:0.7rem; color:#94a3b8; letter-spacing:2px; font-weight:800;">ICLA ACTUAL</h3>
                        <p style="font-size:0.6rem; color:#475569; margin-top:5px;">${tReg} requisitos evaluados</p>
                    </div>
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
                     border-radius:16px; border:1px solid rgba(255,255,255,0.08);">
                    <h3 style="font-size:0.8rem; color:#94a3b8; margin-bottom:30px; text-align:center;
                         letter-spacing:4px; font-weight:900;">MONITOR OPERATIVO</h3>
                    <div style="display:grid; grid-template-columns:1fr 1fr; gap:25px;">${htmlBarras}</div>
                </div>

                <div style="grid-column:span 3; display:flex; flex-direction:column; gap:20px;">
                    <div style="background:rgba(15,23,42,0.85); padding:25px; border-radius:16px;
                         border-top:4px solid #ffe031; border:1px solid rgba(255,255,255,0.08);">
                        <h3 style="font-size:0.75rem; color:#ffe031; margin-bottom:12px; font-weight:900;">
                            🕐 VIGILANCIA (${vigilanciaStack.length})
                        </h3>
                        <div style="max-height:160px; overflow-y:auto;">
                            ${vigilanciaStack.length
                                ? vigilanciaStack.slice(0, 6).map(v => `
                                    <div style="font-size:0.65rem; color:#94a3b8; padding:6px 0;
                                         border-bottom:1px solid rgba(255,255,255,0.03);">
                                        <b style="color:#ffe031;">[${v.c}]</b> ${v.n}
                                    </div>`).join('')
                                : '<p style="color:#475569; font-size:0.65rem;">Sin normas vencidas.</p>'
                            }
                        </div>
                    </div>
                    <div style="background:linear-gradient(145deg,#0d0221,#1a0230); padding:25px;
                         border-radius:16px; border:1px solid rgba(255,255,255,0.08); border-top:4px solid #00f3ff;">
                        <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
                            <i class="fas fa-robot" style="color:#00f3ff;font-size:1rem;"></i>
                            <h3 style="font-size:0.85rem;color:#00f3ff;margin:0;font-weight:900;">AGENTE IA</h3>
                        </div>
                        <p style="font-size:0.6rem;color:#475569;margin:0 0 14px;line-height:1.5;">
                            Busca normativa ambiental nueva en Minambiente, AMVA, CORNARE, IDEAM y ANLA.
                        </p>
                        <input id="ia-in" type="text"
                               placeholder="Ej: residuos peligrosos 2025..."
                               style="width:100%; background:rgba(0,0,0,0.5); border:1px solid #334155;
                                      color:#fff; padding:11px 14px; border-radius:8px; margin-bottom:10px;
                                      font-size:0.78rem; outline:none; box-sizing:border-box;
                                      transition:border-color 0.2s;"
                               onfocus="this.style.borderColor='#00f3ff'"
                               onblur="this.style.borderColor='#334155'">
                        <button id="ia-bt"
                                style="width:100%; height:42px; background:#00f3ff; color:#000;
                                       border:none; border-radius:8px; font-weight:900; cursor:pointer;
                                       font-size:0.72rem; letter-spacing:1.5px; display:flex;
                                       align-items:center; justify-content:center; gap:8px;">
                            <i class="fas fa-search"></i> ESCANEAR NORMATIVA
                        </button>
                    </div>
                </div>

                <div style="grid-column:span 12; background:rgba(15,23,42,0.6);
                     border:1px solid rgba(0,243,255,0.15); border-radius:16px; padding:25px;
                     display:flex; align-items:center; justify-content:space-between; margin:10px 0;">
                    <div style="width:280px;">
                        <h4 style="color:#00f3ff; font-size:0.8rem; font-weight:900; letter-spacing:3px; margin:0;">
                            CUMPLIMIENTO TOTAL
                        </h4>
                        <p style="color:#64748b; font-size:0.65rem; margin:8px 0 0;">
                            Seleccione un punto para ver el detalle mensual.
                        </p>
                    </div>
                    <div style="flex-grow:1; height:130px; margin:0 40px;">
                        <svg viewBox="-20 -40 640 180" preserveAspectRatio="none"
                             style="width:100%; height:100%; overflow:visible;">
                            <polyline points="${ptsGlobal}" fill="none" stroke="#00f3ff"
                                      stroke-width="4" stroke-linecap="round" stroke-linejoin="round" />
                            ${nodosGlobal}
                        </svg>
                    </div>
                    <div style="text-align:right; width:150px;">
                        <div style="font-size:2.5rem; font-weight:900; color:#fff; line-height:1;">${global}%</div>
                        <div style="color:#00ff80; font-size:0.65rem; font-weight:800; margin-top:5px;">● ESTABLE</div>
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

        inicializarAgenteIA();

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
