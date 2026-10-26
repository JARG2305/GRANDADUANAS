require('dotenv').config();
const express = require('express');
const app = express(); 
const path = require('path');
const XLSX = require('xlsx-js-style');
const cron = require('node-cron');
const { default: makeWASocket, useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys');
const qrcode = require('qrcode');
const nodemailer = require('nodemailer');

// ==========================================
// FUNCIÓN CENTRALIZADA PARA GENERAR EL EXCEL
// ==========================================
function generarExcelConsolidadoBuffer(expedientesData) {
    if (!expedientesData || expedientesData.length === 0) return null;

    const excelHeaders = [
        "MODO VÍA", "N.º EXPEDIENTE", "CLIENTE", "PROVEEDOR", "LÍNEA", "PAÍS ORIGEN", "BUQUE / VUELO ORIGEN", "AWB / BL", "N.º CONTENEDOR(ES)", "PESO BL", "CONTENIDO SEGÚN BL", "ETD ORIGEN", "PUERTO TRANSBORDO", "ETA TRANSBORDO", "ETD TRANSBORDO", "BUQUE / VUELO A VE", "ETA LA GUAIRA", "FECHA REGISTRO DUA", "VENCIMIENTO DUA", "FECHA DE LLEGADA", "RECIBIDA ACTA RECEPCIÓN", "FECHA ABANDONO LEGAL", "PERMISOLOGÍA", "FECHA RECIBIDO PERMISOLOGÍA", "PREVALORACIÓN ENVIADA", "DOC. VALORADO EN SISTEMA", "FACTURA RECIBIDA", "MONTO FLETE", "RECIBIDO DOC. TRANSPORTE", "FECHA TRANSMISIÓN", "CANAL", "FUNCIONARIO", "RECONOCIMIENTO", "VALIDACIÓN", "DESPACHO", "ALMACÉN", "RECIBIDA ACTA (ALMACÉN)", "DÍAS LIBRES ALMACÉN", "INICIO DÍAS LIBRES ALMACÉN", "CULMINACIÓN DÍAS LIBRES ALMACÉN", "NAVIERA", "DÍAS LIBRES NAVIERA", "INICIO DÍAS LIBRES NAVIERA", "CULMINACIÓN DÍAS LIBRES NAVIERA", "OBSERVACIONES"
    ];

    const fieldsList = [
        "modo_transporte", "num_expediente", "cliente", "proveedor", "linea", "pais_origen", "buque_origen_vuelo",
        "awb_bl", "numero_contenedor", "peso_bl", "contenido_segun_bl", "etd_origen",
        "puerto_transbordo", "eta_transbordo", "etd_transbordo", "buque_vuelo_ve",
        "eta_la_guaira", "fecha_registro_dai", "vencimiento_dai", "fecha_llegada", "recibida_acta_recepcion", "fecha_abandono_legal", "permisologia", "fecha_recibido_permisologia",
        "prevaloracion_enviada", "documento_valorado_en_sistema", "factura_recibida", "monto_flete",
        "recibido_documento_transporte", "fecha_transmision", "canal", "funcionario", "reconocimiento", "validacion", "despacho",
        "almacen", "recibida_acta_recepcion_almacen", "dias_libres_almacen", "inicio_dias_libres", "culminacion_dias_libres",
        "naviera", "dias_libres_naviera", "inicio_dias_libres_naviera", "culminacion_dias_libres_naviera",
        "observaciones"
    ];

    const porLlegar = expedientesData.filter(e => !e.fecha_llegada && !e.despacho);
    const tramites = expedientesData.filter(e => e.fecha_llegada && !e.despacho);
    const despachados = expedientesData.filter(e => Boolean(e.despacho));

    porLlegar.sort((a, b) => (parseFechaLocalServidor(a.eta_la_guaira) || new Date('9999-12-31')) - (parseFechaLocalServidor(b.eta_la_guaira) || new Date('9999-12-31')));
    tramites.sort((a, b) => (parseFechaLocalServidor(a.eta_la_guaira || a.fecha_llegada) || new Date('9999-12-31')) - (parseFechaLocalServidor(b.eta_la_guaira || b.fecha_llegada) || new Date('9999-12-31')));
    despachados.sort((a, b) => (parseFechaLocalServidor(a.eta_la_guaira || a.fecha_llegada) || new Date('9999-12-31')) - (parseFechaLocalServidor(b.eta_la_guaira || b.fecha_llegada) || new Date('9999-12-31')));

    let wsData = [];
    let metaSecciones = [];

    function agregarSeccion(tituloSeccion, listaItems, colorHex) {
        if (listaItems.length === 0) return;
        wsData.push([tituloSeccion]);
        metaSecciones.push({ type: 'title' });
        wsData.push(excelHeaders);
        metaSecciones.push({ type: 'header', color: colorHex });
        listaItems.forEach(e => {
            let row = fieldsList.map(f => e[f] !== undefined && e[f] !== null ? e[f] : "");
            wsData.push(row);
            metaSecciones.push({ type: 'data' });
        });
        wsData.push([]);
        metaSecciones.push({ type: 'empty' });
    }

    agregarSeccion("1. EXPEDIENTES POR LLEGAR", porLlegar, "FEF08A"); // Amarillo Pollito
    agregarSeccion("2. TRÁMITES ADUANALES Y EN PROCESO", tramites, "2563EB"); // Azul Rey
    agregarSeccion("3. EXPEDIENTES DESPACHADOS", despachados, "DC2626"); // Rojo

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    const sectionTitleStyle = {
        font: { name: "Arial", sz: 12, bold: true, color: { rgb: "1E293B" } },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "left", vertical: "center" }
    };

    const cellStyle = {
        font: { name: "Arial", sz: 10, color: { rgb: "1E293B" } },
        alignment: { horizontal: "left", vertical: "center", wrapText: true },
        border: {
            top: { style: "thin", color: { rgb: "E2E8F0" } },
            bottom: { style: "thin", color: { rgb: "E2E8F0" } },
            left: { style: "thin", color: { rgb: "E2E8F0" } },
            right: { style: "thin", color: { rgb: "E2E8F0" } }
        }
    };

    const range = XLSX.utils.decode_range(ws['!ref'] || "A1");
    for (let R = range.s.r; R <= range.e.r; ++R) {
        const meta = metaSecciones[R] || { type: 'empty' };
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[cellAddress]) continue;

            if (meta.type === 'title') {
                if (C === 0) ws[cellAddress].s = sectionTitleStyle;
            } else if (meta.type === 'header') {
                const textColor = meta.color === "FEF08A" ? "1E293B" : "FFFFFF";
                ws[cellAddress].s = {
                    font: { name: "Arial", sz: 11, bold: true, color: { rgb: textColor } },
                    fill: { fgColor: { rgb: meta.color } },
                    alignment: { horizontal: "center", vertical: "center", wrapText: true },
                    border: {
                        top: { style: "thin", color: { rgb: "CBD5E1" } },
                        bottom: { style: "thin", color: { rgb: "CBD5E1" } },
                        left: { style: "thin", color: { rgb: "CBD5E1" } },
                        right: { style: "thin", color: { rgb: "CBD5E1" } }
                    }
                };
            } else if (meta.type === 'data') {
                ws[cellAddress].s = cellStyle;
            }
        }
    }

    ws['!cols'] = [
        { wch: 11 }, { wch: 9.5 }, { wch: 28 }, { wch: 28 }, { wch: 15 },
        { wch: 12 }, { wch: 23 }, { wch: 23 }, { wch: 28 }, { wch: 11 },
        { wch: 32 }, { wch: 12 }, { wch: 21 }, { wch: 17 }, { wch: 18 },
        { wch: 22 }, { wch: 14 }, { wch: 20 }, { wch: 18 }, { wch: 18 },
        { wch: 24.5 }, { wch: 24 }, { wch: 18 }, { wch: 24 }, { wch: 25 },
        { wch: 24 }, { wch: 18 }, { wch: 13 }, { wch: 26 }, { wch: 20 },
        { wch: 11 }, { wch: 14 }, { wch: 15.5 }, { wch: 12 }, { wch: 12 },
        { wch: 12 }, { wch: 25 }, { wch: 20 }, { wch: 20 }, { wch: 20 },
        { wch: 16 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 35 }
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Reporte Consolidado");
    return XLSX.write(wb, { bookType: 'xlsx', type: 'buffer', cellStyles: true });
}

// ==========================================
// 1. ENVÍO DINÁMICO CON EXCEL ESTILIZADO (VÍA GOOGLE APPS SCRIPT)
// ==========================================
async function enviarCorreoConGoogleScriptDinamico(correoDestino, asuntoCorreo, mensajeCorreo, expedientesData = []) {
    try {
        const urlAppsScript = process.env.googleScriptUrl;
        if (!urlAppsScript) {
            console.error("❌ ERROR CRÍTICO: No se encontró la variable googleScriptUrl en las variables de entorno de Render.");
            return { success: false, message: "URL de Google Apps Script no configurada en Render." };
        }

        let archivoExcelBase64 = null;
        let nombreArchivo = "Reporte_Expedientes_Logistatus.xlsx";

        if (expedientesData && expedientesData.length > 0) {
            const excelBuffer = generarExcelConsolidadoBuffer(expedientesData);
            if (excelBuffer) {
                archivoExcelBase64 = excelBuffer.toString('base64');
            }
        }

        const response = await fetch(urlAppsScript, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                destinatario: correoDestino,
                asunto: asuntoCorreo,
                mensaje: mensajeCorreo,
                excelBase64: archivoExcelBase64,
                nombreArchivo: nombreArchivo
            })
        });

        const textoRespuesta = await response.text();
        let resultado;
        try {
            resultado = JSON.parse(textoRespuesta);
        } catch (e) {
            resultado = { success: false, message: "Error en formato de respuesta de Apps Script: " + textoRespuesta.substring(0, 100) };
        }
        return resultado;
    } catch (error) {
        console.error("❌ Error crítico en enviarCorreoConGoogleScriptDinamico:", error);
        return { success: false, message: error.toString() };
    }
}

// ==========================================
// 2. CONFIGURACIÓN DE MIDDLEWARES
// ==========================================
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.static(path.join(__dirname)));

const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID;
const FIREBASE_PROJECT_ID = process.env.FIREBASE_PROJECT_ID || "statusylogistica";
const FIREBASE_API_KEY = process.env.FIREBASE_API_KEY || "";

app.use((req, res, next) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    if (req.method === 'OPTIONS') return res.sendStatus(200);
    next();
});

// ==========================================
// 3. RUTAS DE LA API Y CONFIGURACIÓN
// ==========================================
app.get('/api/config', (req, res) => {
    res.json({
        firebaseApiKey: FIREBASE_API_KEY,
        firebaseProjectId: FIREBASE_PROJECT_ID,
        telegramBotToken: TELEGRAM_BOT_TOKEN || '',
        telegramChatId: TELEGRAM_CHAT_ID || ''
    });
});

app.get('/', (req, res) => { res.sendFile(path.join(__dirname, 'index.html')); });
app.get('/health', (req, res) => { res.status(200).send('LOGISTATUS PRO - Servidor Activo OK'); });

app.get('/api/clientes', async (req, res) => {
    try {
        const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/clientes`;
        const response = await fetch(url);
        const result = await response.json();
        if (!result.documents) return res.json({ success: true, clientes: [] });

        const clientes = result.documents.map(doc => {
            let fields = doc.fields || {};
            return {
                id: doc.name.split('/').pop(),
                razonSocial: fields.razonSocial?.stringValue || '',
                rif: fields.rif?.stringValue || '',
                direccion: fields.direccion?.stringValue || '',
                telefono: fields.telefono?.stringValue || '',
                correo: fields.correo?.stringValue || '',
                enviarAutomatico: fields.enviarAutomatico?.booleanValue ?? true
            };
        });
        res.json({ success: true, clientes });
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

app.post('/api/clientes', async (req, res) => {
    try {
        const { id, razonSocial, rif, direccion, telefono, correo, enviarAutomatico } = req.body;
        if (!razonSocial) return res.status(400).json({ success: false, message: "La razón social es obligatoria." });

        const isUpdate = Boolean(id);
        const url = isUpdate 
            ? `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/clientes/${id}?updateMask.fieldPaths=razonSocial&updateMask.fieldPaths=rif&updateMask.fieldPaths=direccion&updateMask.fieldPaths=telefono&updateMask.fieldPaths=correo&updateMask.fieldPaths=enviarAutomatico`
            : `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/clientes`;

        const body = {
            fields: {
                razonSocial: { stringValue: razonSocial.trim() },
                rif: { stringValue: (rif || "S/N").trim() },
                direccion: { stringValue: (direccion || "").trim() },
                telefono: { stringValue: (telefono || "").trim() },
                correo: { stringValue: (correo || "").trim() },
                enviarAutomatico: { booleanValue: enviarAutomatico ?? true }
            }
        };

        if (!isUpdate) body.fields.fechaCreacion = { stringValue: new Date().toISOString() };

        const response = await fetch(url, {
            method: isUpdate ? 'PATCH' : 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });

        if (response.ok) res.json({ success: true });
        else {
            const errText = await response.text();
            res.status(500).json({ success: false, message: errText });
        }
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

app.post('/api/forzar-correo-manual', async (req, res) => {
    try {
        const { destinatario, razonSocial, expedientesData } = req.body;
        const excelBuffer = generarExcelConsolidadoBuffer(expedientesData);
        const base64Data = excelBuffer ? excelBuffer.toString('base64') : null;

        const asunto = `Reporte Consolidado de Estatus - Logistatus Pro (Manual)`;
        const mensaje = `Estimado/a ${razonSocial},\n\nAdjunto encontrará el reporte consolidado de estatus solicitado manualmente.\n\nLogistatus Pro`;

        await enviarCorreoConGoogleScriptDinamico(destinatario, asunto, mensaje, expedientesData);
        res.json({ success: true, message: "Correo manual enviado exitosamente." });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

app.post('/api/enviar-correo-cliente', async (req, res) => {
    const { correoCliente, nombreCliente, numExpediente, statusActual, bl, campoLegible, modoTransporte, observaciones, proveedor } = req.body;
    if (!correoCliente) return res.status(400).json({ success: false, message: "El cliente no tiene correo registrado." });

    try {
        const esAereo = String(modoTransporte || "").toLowerCase().includes("aéreo") || String(modoTransporte || "").toLowerCase().includes("aereo");
        const etiquetaDocumento = esAereo ? "Guía (AWB)" : "B/L";
        const asunto = `Actualización de Estatus - Expediente N° ${numExpediente || 'GENERAL'}`;
        const proveedorFinal = proveedor || "S/N";

        let cuerpoCorreo = `
            <div style="font-family: Arial, sans-serif; color: #1e293b; max-width: 500px; margin: 0 auto; padding: 0 8px 4px 8px; border: 1px solid #e2e8f0; border-radius: 8px; background-color: #ffffff;">
                
                <!-- Título y LOGISTATUS PRO pegados absolutamente al borde superior -->
                <div style="border-bottom: 0px solid #2563eb; padding-top: 0px; padding-bottom: 0px; margin-bottom: 3px;">
                    <span style="font-size: 20px; font-weight: 800; color: #1e3a8a; letter-spacing: -2.3px; line-height: 1; display: inline-block; margin: 0;">Actualización de Estatus</span>
                    <span style="font-size: 12px; font-weight: 700; color: #64748b; margin-left: 1px; letter-spacing: 0.5px;">LOGISTATUS PRO</span>
                </div>

                <!-- Saludo pegado arriba -->
                <p style="margin: 0 0 -1px 0; font-size: 14px; color: #334155;">Apreciado cliente <strong>${nombreCliente || 'Estimado'}</strong>, le informamos la actualización reciente de su carga:</p>
                
                <!-- Bloque compacto de datos -->
                <div style="font-size: 14px; line-height: 1.0; background-color: #f8fafc; padding: 0px 0px; border-radius: 6px; margin-bottom: 2px;">
                    <div style="margin: 0px 0;"><strong>N° Exp:</strong> ${numExpediente || 'GENERAL'} &nbsp;|&nbsp; <strong>Proveedor:</strong> ${proveedorFinal} &nbsp;|&nbsp; <strong>${etiquetaDocumento}:</strong> ${bl || 'S/N'}</div>
                    <div style="margin: 1px 0; color: #EF4444;"><strong>${campoLegible || 'Estatus'}:</strong> ${statusActual || 'N/A'}</div>
                    <div style="margin: 1px 0;"><strong>Transporte:</strong> ${modoTransporte || 'Marítimo'}</div>
                    ${observaciones && observaciones.trim() !== "" ? `<div style="margin: 1px 0; color: #475569;"><strong>Observación:</strong> ${observaciones.trim()}</div>` : ''}
                </div>

                <!-- Pie de página reducido -->
                <p style="margin: 1px 0 0 0; font-size: 12px; color: #94a3b8; text-align: center;">Seguimos trabajando para garantizar una gestión eficiente y transparente.</p>
            </div>
        `;

        const resultado = await enviarCorreoConGoogleScriptDinamico(correoCliente, asunto, cuerpoCorreo, []);
        res.json(resultado);
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

  
app.post('/api/enviar-excel-correo', async (req, res) => {
    try {
        const { destinatario, asunto, mensaje, excelBase64, nombreArchivo } = req.body;
        if (!destinatario) return res.status(400).json({ success: false, message: "El destinatario es obligatorio." });

        const urlAppsScript = process.env.googleScriptUrl;
        if (!urlAppsScript) {
            console.error("❌ ERROR: La variable googleScriptUrl no está definida en Render.");
            return res.status(500).json({ success: false, message: "Falta configurar googleScriptUrl en Render." });
        }

        const response = await fetch(urlAppsScript, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ 
                destinatario, 
                asunto, 
                mensaje, 
                excelBase64: excelBase64 || null, 
                nombreArchivo: nombreArchivo || "Reporte_Logistatus.xlsx" 
            })
        });

        const textoRespuesta = await response.text();
        let resultado;
        try {
            resultado = JSON.parse(textoRespuesta);
        } catch (e) {
            return res.status(500).json({ success: false, message: "Google Apps Script devolvió HTML o texto plano en vez de JSON." });
        }

        res.json(resultado);
    } catch (error) {
        console.error("❌ [DEBUG CORREO CRÍTICO] Excepción capturada:", error);
        res.status(500).json({ success: false, message: "Error interno en servidor Node.js: " + error.message });
    }
});

// ==========================================
// 4. TAREAS CRON Y FUNCIONES AUXILIARES
// ==========================================
async function obtenerExpedientesFirestore() {
    try {
        const projectId = process.env.FIREBASE_PROJECT_ID;
        const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/expedientes`;
        const response = await fetch(url);
        const result = await response.json();
        if (!result.documents) return [];

        return result.documents.map(doc => {
            let fields = doc.fields || {};
            let parsedRecord = {};
            for (let key in fields) {
                let valObj = fields[key];
                parsedRecord[key] = valObj.stringValue || valObj.doubleValue || valObj.integerValue || valObj.booleanValue || "";
            }
            return parsedRecord;
        });
    } catch (err) { return []; }
}

async function obtenerClientesFirestore() {
    try {
        const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/clientes`;
        const response = await fetch(url);
        const result = await response.json();
        if (!result.documents) return [];

        return result.documents.map(doc => {
            let fields = doc.fields || {};
            return {
                id: doc.name.split('/').pop(),
                razonSocial: fields.razonSocial?.stringValue || '',
                correo: fields.correo?.stringValue || '',
                enviarAutomatico: fields.enviarAutomatico?.booleanValue ?? true
            };
        });
    } catch (err) { return []; }
}

function parseFechaLocalServidor(str) {
    if (!str) return null;
    str = String(str).trim();
    if (str.includes('/')) {
        const p = str.split('/');
        if (p.length === 3 && p[2].length === 4) return new Date(`${p[2]}-${p[1]}-${p[0]}T00:00:00`);
    } else if (str.includes('-')) {
        const p = str.split('-');
        if (p.length === 3 && p[0].length === 4) return new Date(`${p[0]}-${p[1]}-${p[2]}T00:00:00`);
    }
    return null;
}

// CRON Automático (10:00 AM y 6:00 PM)
cron.schedule('0 10,18 * * *', async () => {
    try {
        const expedientes = await obtenerExpedientesFirestore();
        const clientes = await obtenerClientesFirestore();
        const clientesActivos = clientes.filter(c => c.enviarAutomatico !== false && c.correo);

        for (const cliente of clientesActivos) {
            const clienteNombreBuscado = String(cliente.razonSocial || "").trim().toLowerCase();
            const expedientesCliente = expedientes.filter(e => {
                const clienteExp = String(e.cliente || "").trim().toLowerCase();
                return clienteExp === clienteNombreBuscado && !e.despacho;
            });

            if (expedientesCliente.length > 0) {
                const asunto = `Reporte Automático de Estatus y Expedientes - Logistatus Pro`;
                const mensaje = `Estimado/a ${cliente.razonSocial},\n\nAdjunto encontrará el reporte actualizado con el estatus de sus expedientes vigentes.\n\nLogistatus Pro`;
                await enviarCorreoConGoogleScriptDinamico(cliente.correo, asunto, mensaje, expedientesCliente);
            }
        }
    } catch (error) {
        console.error("❌ Error en CRON automático:", error);
    }
}, { timezone: "America/Caracas" });

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => { console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`); });