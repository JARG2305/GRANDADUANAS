require('dotenv').config();
const express = require('express');
const path = require('path');
const XLSX = require('xlsx');
const cron = require('node-cron');

const app = express();

// ==========================================
// 0. CONFIGURACIÓN DE MIDDLEWARES
// ==========================================
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use(express.static(path.join(__dirname)));

// Credenciales y configuraciones desde variables de entorno
const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID;
const FIREBASE_PROJECT_ID = process.env.FIREBASE_PROJECT_ID || "statusylogistica";
const FIREBASE_API_KEY = process.env.FIREBASE_API_KEY || "";

// ==========================================
// 1. RUTA ESENCIAL DE CONFIGURACIÓN DINÁMICA
// ==========================================
app.get('/api/config', (req, res) => {
    res.json({
        firebaseApiKey: FIREBASE_API_KEY,
        firebaseProjectId: FIREBASE_PROJECT_ID,
        telegramBotToken: TELEGRAM_BOT_TOKEN || '',
        telegramChatId: TELEGRAM_CHAT_ID || ''
    });
});

// ==========================================
// 2. FUNCIONES DE SERVICIO (CORREO, TELEGRAM Y FIRESTORE)
// ==========================================
async function enviarCorreoSistema(opcionesMail) {
    try {
        const apiKey = process.env.RESEND_API_KEY; 
        const remitente = "onboarding@resend.dev";

        if (!apiKey || apiKey.includes("TU_API_KEY")) {
            throw new Error("Falta configurar tu RESEND_API_KEY en el archivo .env o en Render.");
        }

        let destinatarios = opcionesMail.to;
        if (!Array.isArray(destinatarios)) {
            destinatarios = [destinatarios];
        }
        destinatarios = destinatarios.map(e => typeof e === 'string' ? e.trim() : '').filter(e => e.length > 0);

        let attachmentsFormatted = [];
        if (opcionesMail.attachments && Array.isArray(opcionesMail.attachments)) {
            attachmentsFormatted = opcionesMail.attachments.map(att => {
                let base64Content = '';
                if (Buffer.isBuffer(att.content)) {
                    base64Content = att.content.toString('base64');
                } else if (typeof att.content === 'string') {
                    base64Content = att.content.includes('base64,') ? att.content.split('base64,')[1] : att.content;
                }
                return { filename: att.filename, content: base64Content };
            });
        }

        const payload = {
            from: `LOGISTATUS PRO <${remitente}>`,
            to: destinatarios,
            subject: opcionesMail.subject || 'Notificación LOGISTATUS PRO',
            html: opcionesMail.html || opcionesMail.text || 'Sin contenido'
        };

        if (attachmentsFormatted.length > 0) {
            payload.attachments = attachmentsFormatted;
        }

        const response = await fetch('https://api.resend.com/emails', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify(payload)
        });

        const data = await response.json();
        if (!response.ok) {
            throw new Error(`Error de la API: ${JSON.stringify(data)}`);
        }

        console.log("✅ Correo enviado exitosamente vía HTTP:", data.id);
        return { success: true };
    } catch (error) {
        console.error("❌ Error al enviar el correo por HTTP:", error);
        throw error;
    }
}

async function enviarAlertaTelegramServidor(mensajeTexto) {
    try {
        const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                chat_id: TELEGRAM_CHAT_ID,
                text: mensajeTexto,
                parse_mode: 'Markdown'
            })
        });
        const data = await response.json();
        if (!data.ok) {
            console.error("❌ Error de Telegram en el servidor:", data.description);
        } else {
            console.log("✅ Alerta enviada a Telegram desde la nube.");
        }
    } catch (err) {
        console.error("❌ Error de red al conectar con Telegram desde el servidor:", err);
    }
}

// Función para obtener los expedientes directo desde Firebase Firestore vía REST API
async function obtenerExpedientesFirestore() {
    try {
        const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/expedientes`;
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
    } catch (err) {
        console.error("❌ Error al consultar Firestore desde el servidor:", err);
        return [];
    }
}

// Funciones auxiliares de fecha y modo de transporte para el servidor
function parseFechaLocalServidor(str) {
    if (!str) return null;
    str = String(str).trim();
    if (str.includes('/')) {
        const p = str.split('/');
        if (p.length === 3 && p[2].length === 4) {
            return new Date(`${p[2]}-${p[1]}-${p[0]}T00:00:00`);
        }
    } else if (str.includes('-')) {
        const p = str.split('-');
        if (p.length === 3 && p[0].length === 4) {
            return new Date(`${p[0]}-${p[1]}-${p[2]}T00:00:00`);
        }
    }
    return null;
}

function identificarEsMaritimoServidor(rec = {}) {
    const modo = (rec.modo_transporte || "").toLowerCase();
    if (modo.includes("aéreo") || modo.includes("aereo")) return false;
    return true;
}

// ==========================================
// 3. RUTAS DE LA API
// ==========================================
app.post('/api/enviar-excel-correo', async (req, res) => {
    const { destinatario, asunto, mensaje, excelBase64, nombreArchivo } = req.body;
    
    if (!excelBase64) {
        return res.status(400).json({ success: false, message: 'Falta el archivo Excel en Base64.' });
    }

    const base64Data = excelBase64.includes('base64,') ? excelBase64.split('base64,')[1] : excelBase64;

    const mailOptions = {
        to: destinatario,
        subject: asunto || '📊 Reporte de Expedientes - LOGISTATUS PRO',
        text: mensaje || 'Adjunto encontrarás el reporte de expedientes actualizado.',
        attachments: [
            {
                filename: nombreArchivo || 'Reporte_Logistatus.xlsx',
                content: Buffer.from(base64Data, 'base64'),
                contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            }
        ]
    };

    try {
        await enviarCorreoSistema(mailOptions);
        res.json({ success: true, message: 'Correo con Excel adjunto enviado con éxito.' });
    } catch (error) {
        console.error('Error al enviar el correo:', error);
        res.status(500).json({ success: false, message: error.message || 'Error al enviar el correo.' });
    }
});

app.post('/api/notificar', async (req, res) => {
    try {
        const { mensaje } = req.body;
        if (!mensaje) {
            return res.status(400).json({ success: false, message: 'Falta el mensaje de alerta.' });
        }
        await enviarAlertaTelegramServidor(mensaje);
        res.json({ success: true, message: "Alerta enviada a Telegram con éxito." });
    } catch (error) {
        console.error("❌ Error en /api/notificar:", error);
        res.status(500).json({ success: false, message: error.message });
    }
});

// ==========================================
// 4. TAREAS PROGRAMADAS EN SEGUNDO PLANO (CRON)
// ==========================================

// Envío de correos automáticos programados (Ajustado a hora Venezuela: 10:00 AM y 6:00 PM)
cron.schedule('0 10,18 * * *', async () => {
    console.log("⏰ [CRON] Ejecutando envío programado de reporte...");
    try {
        const destinatariosAutomaticos = ["importacionesepga@gmail.com", "hnoguera@gmail.com", "finanzasepga@gmail.com"];
        const mailOptionsAuto = {
            to: destinatariosAutomaticos,
            subject: "Reporte Maestro Unificado de Expedientes - LOGISTATUS PRO",
            text: "Este es el envío automático programado del reporte consolidado de operaciones.",
            html: "<p>Este es el envío automático programado del reporte consolidado de operaciones de <strong>Logistatus Pro</strong>.</p>"
        };
        await enviarCorreoSistema(mailOptionsAuto);
    } catch (error) {
        console.error("❌ [CRON] Error al enviar el correo automático programado:", error);
    }
}, {
    timezone: "America/Caracas"
});

// Evaluación autónoma de alertas operativas en hora de Venezuela (9 AM, 12 PM, 3 PM, 6 PM, 7 PM, 8 PM)
cron.schedule('0 9,12,15,18,19,20 * * *', async () => {
    const ahoraVzla = new Date().toLocaleString("en-US", { timeZone: "America/Caracas" });
    const horaActual = new Date(ahoraVzla).getHours();
    console.log(`⏰ [CRON NUBE] Evaluando alertas operativas en hora Venezuela (${horaActual}:00 hrs)...`);
    
    try {
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) return;

        const today = new Date();
        today.setHours(0,0,0,0);

        let expSet = new Set();
        let blSet = new Set();
        let contSet = new Set();
        let alertasEnviadasCount = 0;

        for (const rec of records) {
            const expVal = (rec.num_expediente || "").trim().toLowerCase();
            const blVal = (rec.awb_bl || "").trim().toLowerCase();
            const contRaw = (rec.numero_contenedor || "").trim();
            let contenedoresList = contRaw ? contRaw.split(',').map(c => c.trim().toLowerCase()).filter(Boolean) : [];

            if (expVal) {
                if (expSet.has(expVal)) {
                    await enviarAlertaTelegramServidor(`❗ **DUPLICADO:** El N.º de Expediente **${rec.num_expediente}** está repetido en el sistema.`);
                    alertasEnviadasCount++;
                }
                expSet.add(expVal);
            }
            if (blVal && blVal !== 's/n') {
                if (blSet.has(blVal)) {
                    await enviarAlertaTelegramServidor(`❗ **DUPLICADO:** El N.º de B/L o AWB **${rec.awb_bl}** ya está usado en otro expediente.`);
                    alertasEnviadasCount++;
                }
                blSet.add(blVal);
            }
            contenedoresList.forEach(async (contVal) => {
                if (contSet.has(contVal)) {
                    await enviarAlertaTelegramServidor(`❗ **DUPLICADO:** El Contenedor **${contVal.toUpperCase()}** ya está registrado en otro expediente.`);
                    alertasEnviadasCount++;
                }
                contSet.add(contVal);
            });

            if (rec.despacho) continue;

            const expName = rec.num_expediente || 'S/N';
            const clientName = rec.cliente || 'S/C';
            const esMaritimo = identificarEsMaritimoServidor(rec);
            const fechaLlegadaStr = rec.fecha_llegada;
            const hasLlegadaRegistrada = Boolean(fechaLlegadaStr && fechaLlegadaStr.trim() !== "");

            const fechaArriboObjetivo = rec.fecha_llegada || rec.eta_la_guaira;
            const dArriboObj = parseFechaLocalServidor(fechaArriboObjetivo);

            if (!hasLlegadaRegistrada) {
                const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
                if (dVencimientoDai) {
                    const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                    if (diffVencDays >= 1 && diffVencDays <= 4) {
                        let textoCountdown = diffVencDays === 1 ? "¡VENCE MAÑANA!" : `Faltan ${diffVencDays} día(s) para vencer`;
                        const textTelegram = `**⚠️ DAI PRÓXIMA A VENCER - ${textoCountdown}**\n\n📋 *Expediente:* **${expName}**\n👤 *Cliente:* **${clientName}**\n⏳ Vencimiento: *${rec.vencimiento_dai}*.`;
                        
                        await enviarAlertaTelegramServidor(textTelegram);
                        alertasEnviadasCount++;
                    }
                }

                if (dArriboObj) {
                    const diffLlegadaDays = Math.round((dArriboObj - today) / (1000 * 60 * 60 * 24));
                    const maxDiasAnticipacion = esMaritimo ? 5 : 4;
                    const minDiasAnticipacion = esMaritimo ? 2 : 1;

                    if (diffLlegadaDays >= minDiasAnticipacion && diffLlegadaDays <= maxDiasAnticipacion) {
                        const modoTexto = esMaritimo ? "MARÍTIMO" : "AÉREO";
                        const tipoFecha = rec.fecha_llegada ? "Llegada" : "ETA La Guaira";
                        const textTelegram = `**⚠️ CUENTA REGRESIVA ${tipoFecha.toUpperCase()} (${modoTexto})**\n\n📋 *Expediente:* **${expName}**\n👤 *Cliente:* **${clientName}**\n⏳ Faltan *${diffLlegadaDays} día(s)* para el ${tipoFecha} (${fechaArriboObjetivo}).`;

                        await enviarAlertaTelegramServidor(textTelegram);
                        alertasEnviadasCount++;
                    }

                    const noTieneDai = !rec.fecha_registro_dai || rec.fecha_registro_dai.trim() === "";
                    if (diffLlegadaDays <= maxDiasAnticipacion && noTieneDai) {
                        const textTelegramDai = `**⚠️ FALTA REGISTRAR DAI**\n\n📋 *Expediente:* **${expName}**\n👤 *Cliente:* **${clientName}**\n⏳ El arribo es cercano (*${fechaArriboObjetivo}*) y el registro DAI está pendiente o vacío.`;

                        await enviarAlertaTelegramServidor(textTelegramDai);
                        alertasEnviadasCount++;
                    }
                }
            }
        }
        console.log(`✅ [CRON NUBE] Verificación finalizada. Alertas enviadas: ${alertasEnviadasCount}`);
    } catch (error) {
        console.error("❌ [CRON NUBE] Error al procesar alertas automáticas:", error);
    }
}, {
    timezone: "America/Caracas"
});

// ==========================================
// 5. INICIO DEL SERVIDOR
// ==========================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`);
});