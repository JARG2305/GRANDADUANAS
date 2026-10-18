require('dotenv').config();
const express = require('express');
const app = express(); 
const path = require('path');
const XLSX = require('xlsx-js-style');
const cron = require('node-cron');
const { default: makeWASocket, useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys');
const qrcode = require('qrcode');

// ==========================================
// 0. CONFIGURACIÓN DE MIDDLEWARES
// ==========================================
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
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
// 0.2 INICIALIZACIÓN DE WHATSAPP (BAILEYS)
// ==========================================
let sock = null;
let whatsappClientReady = false;
let ultimoQrGenerado = "";
let estadoConexionWhatsapp = "DESCONECTADO";

async function iniciarBaileysWhatsApp() {
    try {
        const { state, saveCreds } = await useMultiFileAuthState(path.join(__dirname, 'auth_info_baileys'));
        
        sock = makeWASocket({
            auth: state,
            printQRInTerminal: false
        });

        sock.ev.on('connection.update', async (update) => {
            const { connection, lastDisconnect, qr } = update;

            if (qr) {
                estadoConexionWhatsapp = "ESPERANDO_ESCANEO";
                try {
                    ultimoQrGenerado = await qrcode.toDataURL(qr);
                } catch (err) { console.error("Error QR:", err); }
            }

            if (connection === 'open') {
                whatsappClientReady = true;
                estadoConexionWhatsapp = "CONECTADO";
                ultimoQrGenerado = "";
            }

            if (connection === 'close') {
                whatsappClientReady = false;
                estadoConexionWhatsapp = "DESCONECTADO";
                const shouldReconnect = lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut;
                if (shouldReconnect) setTimeout(iniciarBaileysWhatsApp, 5000);
            }
        });

        sock.ev.on('creds.update', saveCreds);

        // ESCUCHA DE MENSAJES EN WHATSAPP (COMANDOS)
        sock.ev.on('messages.upsert', async (m) => {
            try {
                if (m.type !== 'notify') return;
                const msg = m.messages[0];
                if (!msg.message || msg.key.fromMe) return;

                const remitente = msg.key.remoteJid;
                const textoMensaje = (msg.message.conversation || msg.message.extendedTextMessage?.text || "").trim();
                const textoLower = textoMensaje.toLowerCase();
                const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us"; 
                
                if (remitente !== GRUPO_AUTORIZADO_JID) return;

                if (textoLower.includes('resumen') || textoLower.includes('conteo') || textoLower.includes('totales')) {
                    const records = await obtenerExpedientesFirestore();
                    if (!records || records.length === 0) {
                        await sock.sendMessage(remitente, { text: "✅ No hay registros." });
                        return;
                    }
                    await sock.sendMessage(remitente, { text: `📊 *RESUMEN DE EXPEDIENTES*\nTotal registros activos: ${records.filter(r => !r.despacho).length}` });
                    return;
                }
            } catch (err) { console.error("Error en Baileys upsert:", err); }
        });
    } catch (err) { console.error("Error al iniciar Baileys:", err); }
}
iniciarBaileysWhatsApp();

// ==========================================
// 0.3 INICIALIZACIÓN DE TELEGRAM (COMANDOS Y POLLING)
// ==========================================
async function iniciarTelegramPolling() {
    if (!TELEGRAM_BOT_TOKEN) return;
    let offset = 0;
    
    setInterval(async () => {
        try {
            const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/getUpdates?offset=${offset}&timeout=30`;
            const response = await fetch(url);
            const data = await response.json();

            if (data.ok && data.result.length > 0) {
                for (const update of data.result) {
                    offset = update.update_id + 1;
                    if (!update.message || !update.message.text) continue;

                    const chatId = String(update.message.chat.id);
                    const texto = update.message.text.trim().toLowerCase();

                    // Validar si el chat coincide con el autorizado
                    if (TELEGRAM_CHAT_ID && chatId !== String(TELEGRAM_CHAT_ID)) continue;

                    if (texto.includes('resumen') || texto.includes('conteo') || texto.includes('alertas')) {
                        const records = await obtenerExpedientesFirestore();
                        let msgRespuesta = `📊 *REPORTE SOLICITADO DESDE TELEGRAM*\nTotal registros activos: ${records.filter(r => !r.despacho).length}`;
                        await enviarAlertaTelegramServidor(msgRespuesta);
                    }
                }
            }
        } catch (err) {
            // Ignorar errores de red temporales en polling
        }
    }, 5000);
}
iniciarTelegramPolling();

// ==========================================
// 1. RUTAS ESENCIALES
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

app.get('/api/whatsapp/estado', (req, res) => {
    res.json({ success: true, estado: estadoConexionWhatsapp, ready: whatsappClientReady, qrCodeDataUrl: ultimoQrGenerado });
});

// =========================================================================
// RUTA DE FORZAR ALERTAS (RESPETA ESTRICTAMENTE LOS CANALES ENCENDIDOS)
// =========================================================================
app.post('/api/alertas/forzar', async (req, res) => {
    try {
        const { canales } = req.body; // { whatsapp: true/false, telegram: true/false }
        
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) {
            return res.json({ success: false, message: "No hay registros en Firestore para evaluar." });
        }

        let alertasEta = [], alertasVencimientos = [];
        const today = new Date();
        today.setHours(0,0,0,0);

        for (const rec of records) {
            if (rec.despacho) continue;

            const expName = rec.num_expediente || 'S/N';
            const clientName = rec.cliente || 'S/C';
            const blInfo = rec.awb_bl || rec.bl || rec.numero_bl || 'S/N'; 
            const fechaEtaRaw = rec.fecha_llegada || rec.eta_la_guaira || 'S/N';
            const fechaEtaFormatted = formatearFechaDDMMYYYY(fechaEtaRaw);
            
            const modoBruto = (rec.modo_transporte || "").toLowerCase().trim();
            const esBarco = !modoBruto.includes("aéreo") && !modoBruto.includes("aereo");
            const iconoVia = esBarco ? "🚢" : "✈️";
            const destinoTxt = esBarco ? "ETA La Guaira" : "ETA Maiquetía";

            const hasLlegadaRegistrada = Boolean(rec.fecha_llegada && rec.fecha_llegada.trim() !== "");
            const hasRegistroDai = Boolean(rec.fecha_registro_dai && rec.fecha_registro_dai.trim() !== "");
            const hasTransmision = Boolean(rec.fecha_transmision && rec.fecha_transmision.trim() !== "");
            const dArriboObj = parseFechaLocalServidor(fechaEtaRaw);

            // 1. Alertas de ETA
            if (dArriboObj && !hasLlegadaRegistrada) {
                const diffLlegadaDays = Math.round((dArriboObj - today) / (1000 * 60 * 60 * 24));
                if (diffLlegadaDays >= 0 && diffLlegadaDays <= 5) {
                    const textoDias = diffLlegadaDays === 0 ? "¡LLEGA HOY!" : `Faltan ${diffLlegadaDays}d`;
                    let detalleDaiTexto = hasRegistroDai ? `\n📄 DAI Registrada` : `\n⚠️ *FALTA REGISTRAR DAI / DUA*`;
                    alertasEta.push(`${iconoVia} *${destinoTxt.toUpperCase()}*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n⏳ *${textoDias}* (${fechaEtaFormatted})${detalleDaiTexto}`);
                }
            }

            // 2. Plazo legal Art. 47 LOA
            const dLlegada = parseFechaLocalServidor(rec.fecha_llegada);
            if (dLlegada && !hasTransmision) {
                let dDeadline = new Date(dLlegada.getTime());
                let diasHabiles = 0, intentos = 0;
                while (diasHabiles < 5 && intentos < 30) {
                    dDeadline.setDate(dDeadline.getDate() + 1);
                    if (dDeadline.getDay() !== 0 && dDeadline.getDay() !== 6) diasHabiles++;
                    intentos++;
                }
                const diffPlazoDays = Math.round((dDeadline - today) / (1000 * 60 * 60 * 24));
                const fechaPlazoStr = formatearFechaDDMMYYYY(dDeadline.toISOString().split('T')[0]);

                if (diffPlazoDays < 0) {
                    alertasVencimientos.push(`🚨 *EXTEMPORANEIDAD - ART. 47*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n⏳ Venció el ${fechaPlazoStr} (${Math.abs(diffPlazoDays)}d en extemporaneidad).\n🚨 *¡Aplica Multa Art. 177 Num 7!*`);
                }
            }

            // 3. Vencimiento Almacén
            if (rec.culminacion_dias_libres) {
                const culmDate = parseFechaLocalServidor(rec.culminacion_dias_libres);
                if (culmDate) {
                    const diffAlm = Math.round((culmDate - today) / (1000 * 60 * 60 * 24));
                    if (diffAlm < 0) {
                        alertasVencimientos.push(`🚨 *ALMACÉN VENCIDO*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n⚠️ Venció hace ${Math.abs(diffAlm)} día(s)`);
                    }
                }
            }

            // 4. Vencimiento DUA / DAI (Art. 177 Num 9) - Captura Soto Aduanas
            const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
            if (!hasLlegadaRegistrada && !hasRegistroDai && dVencimientoDai) {
                const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                if (diffVencDays < 0) {
                    const diasVencida = Math.abs(diffVencDays);
                    alertasVencimientos.push(`🚨 *VENCIDO / MULTA ART. 177*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n⚠️ *Venció el ${rec.vencimiento_dai}* (${diasVencida}d atrás)`);
                }
            }
        }

        let listaFinalOrdenada = [...alertasEta, ...alertasVencimientos];
        if (listaFinalOrdenada.length === 0) {
            return res.json({ success: false, message: "No hay alertas activas en este momento para enviar." });
        }

        // VALIDACIÓN ESTRICTA: Solo envía a los canales que el usuario tenga encendidos en la interfaz
        let canalesActivos = {
            whatsapp: Boolean(canales?.whatsapp && whatsappClientReady && sock),
            telegram: Boolean(canales?.telegram && TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID)
        };

        await enviarAlertaMulticanal(listaFinalOrdenada, "REPORTE FORZADO MANUAL", canalesActivos);
        res.json({ 
            success: true, 
            message: "Alertas procesadas y enviadas con éxito a las apps encendidas.",
            canalesUtilizados: canalesActivos
        });

    } catch (error) {
        console.error("❌ Error al forzar alertas:", error);
        res.status(500).json({ success: false, message: error.message });
    }
});

// ==========================================
// 2. FUNCIONES AUXILIARES
// ==========================================
function formatearFechaDDMMYYYY(str) {
    const d = parseFechaLocalServidor(str);
    if (!d || isNaN(d.getTime())) return str || 'S/N';
    return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
}

async function enviarAlertaMulticanal(listaAlertas, tituloReporte, canalesActivos) {
    try {
        if (canalesActivos.telegram && TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID) {
            let mensajesEnumerados = listaAlertas.map((item, index) => `${index + 1}️⃣ ${item}`);
            let mensajeConsolidado = `🚨 *LOGISTATUS PRO - ${tituloReporte}* 🚨\n\n` + mensajesEnumerados.join("\n\n. . . . . . . . . . . . . . . . . . . . . . . .\n\n");
            await enviarAlertaTelegramServidor(mensajeConsolidado);
        }

        if (canalesActivos.whatsapp && whatsappClientReady && sock) {
            const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us";
            let mensajesEnumerados = listaAlertas.map((item, index) => `${index + 1}️⃣ ${item}`);
            let mensajeConsolidado = `🚨 *LOGISTATUS PRO - ${tituloReporte}* 🚨\n\n` + mensajesEnumerados.join("\n\n. . . . . . . . . . . . . . . . . . . . . . . .\n\n");
            await sock.sendMessage(GRUPO_AUTORIZADO_JID, { text: mensajeConsolidado });
        }
    } catch (err) {
        console.error("❌ Error en envío multicanal:", err);
    }
}

async function enviarAlertaTelegramServidor(mensajeTexto) {
    try {
        const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
        await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ chat_id: TELEGRAM_CHAT_ID, text: mensajeTexto, parse_mode: 'Markdown' })
        });
    } catch (err) { console.error("Error Telegram:", err); }
}

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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => { console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`); });