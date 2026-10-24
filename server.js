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
// ENVÍO DINÁMICO CON EXCEL ESTILIZADO (VÍA GOOGLE APPS SCRIPT)
// ==========================================
async function enviarCorreoConGoogleScriptDinamico(correoDestino, asuntoCorreo, mensajeCorreo, expedientesData = []) {
    try {
        const urlAppsScript = process.env.googleScriptUrl;
        if (!urlAppsScript) {
            console.error("❌ No se encontró la variable googleScriptUrl en las variables de entorno de Render.");
            return { success: false, message: "URL de Google Apps Script no configurada." };
        }

        let archivoExcelBase64 = null;
        let nombreArchivo = "Reporte_Expedientes_Logistatus.xlsx";

        if (expedientesData && expedientesData.length > 0) {
            const wsData = [
                ["N° Expediente", "Cliente", "Proveedor", "B/L o Guía", "ETA La Guaira", "Fecha Llegada", "Canal", "Estatus / Observaciones"]
            ];

            expedientesData.forEach(e => {
                wsData.push([
                    e.num_expediente || 'S/N',
                    e.cliente || 'S/C',
                    e.proveedor || 'N/A',
                    e.awb_bl || e.bl || 'S/N',
                    e.eta_la_guaira || 'S/N',
                    e.fecha_llegada || 'Pendiente',
                    e.canal || 'Sin asignar',
                    e.observaciones || ''
                ]);
            });

            const ws = XLSX.utils.aoa_to_sheet(wsData);

            const headerStyle = {
                font: { name: "Arial", sz: 11, bold: true, color: { rgb: "FFFFFF" } },
                fill: { fgColor: { rgb: "2563EB" } },
                alignment: { horizontal: "center", vertical: "center" },
                border: {
                    top: { style: "thin", color: { rgb: "CBD5E1" } },
                    bottom: { style: "thin", color: { rgb: "CBD5E1" } },
                    left: { style: "thin", color: { rgb: "CBD5E1" } },
                    right: { style: "thin", color: { rgb: "CBD5E1" } }
                }
            };

            const cellStyle = {
                font: { name: "Arial", sz: 10, color: { rgb: "1E293B" } },
                alignment: { horizontal: "left", vertical: "center" },
                border: {
                    top: { style: "thin", color: { rgb: "E2E8F0" } },
                    bottom: { style: "thin", color: { rgb: "E2E8F0" } },
                    left: { style: "thin", color: { rgb: "E2E8F0" } },
                    right: { style: "thin", color: { rgb: "E2E8F0" } }
                }
            };

            const range = XLSX.utils.decode_range(ws['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!ws[cellAddress]) continue;
                    
                    if (R === 0) {
                        ws[cellAddress].s = headerStyle;
                    } else {
                        ws[cellAddress].s = cellStyle;
                    }
                }
            }

            ws['!cols'] = [
                { wch: 18 }, { wch: 25 }, { wch: 20 }, { wch: 20 },
                { wch: 15 }, { wch: 15 }, { wch: 12 }, { wch: 35 }
            ];

            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Expedientes");

            const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer', cellStyles: true });
            archivoExcelBase64 = excelBuffer.toString('base64');
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

        const resultado = await response.json();
        console.log(`📧 [APPS SCRIPT CON EXCEL ESTILIZADO] Correo enviado a ${correoDestino}`);
        return resultado;

    } catch (error) {
        console.error("❌ Error al enviar Excel estilizado por Google Apps Script:", error);
        return { success: false, message: error.message };
    }
}

// ==========================================
// 0. CONFIGURACIÓN DE MIDDLEWARES
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
    if (req.method === 'OPTIONS') {
        return res.sendStatus(200);
    }
    next();
});

// ==========================================
// 1. RUTAS ESENCIALES Y DE CONFIGURACIÓN
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

app.post('/api/enviar-correo-cliente', async (req, res) => {
    const { correoCliente, numExpediente, bl, eventoCambiado, observaciones } = req.body;
    if (!correoCliente) return res.status(400).json({ success: false, message: "El cliente no tiene correo." });

    try {
        const asunto = `Actualización de Expediente N° ${numExpediente || 'XXXX'}`;
        let mensaje = `Le informamos que su expediente\nN° ${numExpediente || 'XXXX'}\nB/L o Guía: ${bl || 'XXXXX'}`;
        if (eventoCambiado) mensaje += `\n${eventoCambiado}`;
        if (observaciones && observaciones.trim() !== "") mensaje += `\n${observaciones}`;
        mensaje += `\n\nEste es un mensaje automatizado generado por Logistatus Pro.`;

        // 🔍 Buscamos el expediente para adjuntar su Excel estilizado
        const expedientes = await obtenerExpedientesFirestore();
        const expedienteEncontrado = expedientes.find(e => String(e.num_expediente).trim() === String(numExpediente).trim());
        const datosAdjuntos = expedienteEncontrado ? [expedienteEncontrado] : [];

        // 🚀 Se envía con los datos adjuntos para que se cree el Blob en Apps Script
        const resultado = await enviarCorreoConGoogleScriptDinamico(correoCliente, asunto, mensaje, datosAdjuntos);
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
        const response = await fetch(urlAppsScript, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ destinatario, asunto, mensaje, excelBase64, nombreArchivo })
        });

        const resultado = await response.json();
        res.json(resultado);
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => { console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`); });