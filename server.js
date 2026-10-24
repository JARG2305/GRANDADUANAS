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
// 1. ENVÍO DINÁMICO CON EXCEL ESTILIZADO (VÍA GOOGLE APPS SCRIPT)
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
                fill: { fgColor: { rgb: "2563EB" } }, // Azul institucional
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
    if (req.method === 'OPTIONS') {
        return res.sendStatus(200);
    }
    next();
});

// ==========================================
// 3. INICIALIZACIÓN DE WHATSAPP CON BAILEYS
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
                console.log('📱 [WHATSAPP] Nuevo código QR generado.');
                estadoConexionWhatsapp = "ESPERANDO_ESCANEO";
                try {
                    ultimoQrGenerado = await qrcode.toDataURL(qr);
                } catch (err) {
                    console.error("❌ Error al convertir QR:", err);
                }
            }

            if (connection === 'open') {
                console.log('✅ [WHATSAPP] ¡Conectado y listo!');
                whatsappClientReady = true;
                estadoConexionWhatsapp = "CONECTADO";
                ultimoQrGenerado = "";

                try {
                    const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us";
                    await sock.sendMessage(GRUPO_AUTORIZADO_JID, { text: "🟢 *El bot de aduanas se ha conectado y está operativo.*" });
                } catch (e) { console.error("No se pudo enviar aviso de conexión", e); }
            }

            if (connection === 'close') {
                const statusCode = lastDisconnect?.error?.output?.statusCode;
                const isLoggedOut = statusCode === DisconnectReason.loggedOut;
                const shouldReconnect = !isLoggedOut;
                
                whatsappClientReady = false;
                estadoConexionWhatsapp = "DESCONECTADO";

                if (shouldReconnect) {
                    setTimeout(iniciarBaileysWhatsApp, 5000);
                }
            }
        });

        sock.ev.on('creds.update', saveCreds);

        sock.ev.on('messages.upsert', async (m) => {
            try {
                if (m.type !== 'notify') return;
                const msg = m.messages[0];
                if (!msg.message || msg.key.fromMe) return;

                const remitente = msg.key.remoteJid;
                const textoMensaje = (msg.message.conversation || msg.message.extendedTextMessage?.text || "").trim();
                const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us"; 
                if (remitente !== GRUPO_AUTORIZADO_JID) return;

                const respuestaProcesada = await procesarComandoLogistica(textoMensaje);
                if (respuestaProcesada) {
                    await sock.sendMessage(remitente, { text: respuestaProcesada });
                }
            } catch (err) {
                console.error("❌ Error en Baileys upsert:", err);
            }
        });
    } catch (err) {
        console.error("❌ Error al iniciar Baileys:", err);
    }
}
iniciarBaileysWhatsApp();

// ==========================================
// 4. INICIALIZACIÓN DE TELEGRAM (POLLING)
// ==========================================
async function iniciarTelegramPolling() {
    if (!TELEGRAM_BOT_TOKEN) return;
    let offset = 0;
    
    console.log("🤖 [TELEGRAM] Iniciando sondeo (Polling) de comandos...");

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
                    const textoMensaje = update.message.text.trim();

                    if (TELEGRAM_CHAT_ID && chatId !== String(TELEGRAM_CHAT_ID)) continue;

                    const respuestaProcesada = await procesarComandoLogistica(textoMensaje);
                    if (respuestaProcesada) {
                        await enviarAlertaTelegramServidorPersonalizado(chatId, respuestaProcesada);
                    }
                }
            }
        } catch (err) {}
    }, 4000);
}
iniciarTelegramPolling();

async function enviarAlertaTelegramServidorPersonalizado(chatIdDestino, mensajeTexto) {
    if (!chatIdDestino) return;
    try {
        const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
        await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ chat_id: chatIdDestino, text: mensajeTexto, parse_mode: 'Markdown' })
        });
    } catch (err) {
        console.error("❌ Error enviando respuesta en Telegram:", err);
    }
}

// ==========================================
// 5. MOTOR CENTRAL DE COMANDOS Y ALERTAS (CON ART. 47 Y ART. 177)
// ==========================================
function sumarDiasHabiles(fechaInicio, dias) {
    let d = new Date(fechaInicio);
    let diasAgregados = 0;
    while (diasAgregados < dias) {
        d.setDate(d.getDate() + 1);
        let diaSemana = d.getDay();
        if (diaSemana !== 0 && diaSemana !== 6) { // 0 = Domingo, 6 = Sábado
            diasAgregados++;
        }
    }
    return d;
}

async function generarListaAlertasSistema() {
    const records = await obtenerExpedientesFirestore();
    if (!records || records.length === 0) return [];

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
        const dArriboObj = parseFechaLocalServidor(fechaEtaRaw);

        // 1. Alertas de ETA próximo
        if (dArriboObj && !hasLlegadaRegistrada) {
            const diffLlegadaDays = Math.round((dArriboObj - today) / (1000 * 60 * 60 * 24));
            if (diffLlegadaDays >= 0 && diffLlegadaDays <= 5) {
                const textoDias = diffLlegadaDays === 0 ? "¡LLEGA HOY!" : `Faltan ${diffLlegadaDays}d`;
                
                let detalleDaiTexto = "";
                if (hasRegistroDai) {
                    const fechaDaiFormatted = formatearFechaDDMMYYYY(rec.fecha_registro_dai);
                    const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
                    if (dVencimientoDai) {
                        const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                        detalleDaiTexto = diffVencDays < 0 ? `\n📄 DAI (${fechaDaiFormatted}) - ⚠️ VIGENCIA VENCIDA (Art. 177 Num. 9)` : `\n📄 DAI (${fechaDaiFormatted}) - ✅ Vigente`;
                    } else {
                        detalleDaiTexto = `\n📄 DAI Registrada (${fechaDaiFormatted})`;
                    }
                } else {
                    detalleDaiTexto = `\n⚠️ *FALTA REGISTRAR DAI*`;
                }

                alertasEta.push({
                    texto: `${iconoVia} *${destinoTxt.toUpperCase()}*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n⏳ *${textoDias}* (${fechaEtaFormatted})${detalleDaiTexto}`,
                    fechaSort: dArriboObj
                });
            }
        }

        // 2. Control del Art. 47 (5 días hábiles tras la llegada) y Multa Art. 177 Num. 7
        const dLlegadaObj = parseFechaLocalServidor(rec.fecha_llegada);
        const hasTransmision = Boolean(rec.fecha_transmision && rec.fecha_transmision.trim() !== "");
        
        if (dLlegadaObj && !hasTransmision) {
            const fechaLimiteArt47 = sumarDiasHabiles(dLlegadaObj, 5);
            fechaLimiteArt47.setHours(0,0,0,0);
            const diffDiasPlazo = Math.round((fechaLimiteArt47 - today) / (1000 * 60 * 60 * 24));

            if (diffDiasPlazo < 0) {
                const diasRetraso = Math.abs(diffDiasPlazo);
                alertasVencimientos.push({
                    texto: `🚨 *VENCIDO / MULTA ART. 177 NUM. 7 DE LA LOA (ART. 47 LOA)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⚠️ *Plazo de 5 días hábiles vencido* (${diasRetraso}d de retraso)`,
                    fechaSort: fechaLimiteArt47
                });
            } else if (diffDiasPlazo <= 2) {
                const textoPlazo = diffDiasPlazo === 0 ? "¡VENCE HOY EL PLAZO HÁBIL!" : `Vence en ${diffDiasPlazo}d hábil(es)`;
                alertasEta.push({
                    texto: `⏰ *PLAZO ART. 47 (5 DÍAS HÁBILES)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n⏳ *${textoPlazo}* (Límite: ${formatearFechaDDMMYYYY(fechaLimiteArt47)})`,
                    fechaSort: fechaLimiteArt47
                });
            }
        }

        // 3. Control general de vencimiento de DAI (Art. 177 Num. 9)
        const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
        if (dVencimientoDai) {
            const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
            if (diffVencDays < 0) {
                const diasVencida = Math.abs(diffVencDays);
                alertasVencimientos.push({
                    texto: `🚨 *VENCIDO / MULTA ART. 177 NUM. 9 DE LA LOA (DAI)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⚠️ *La DAI venció el ${rec.vencimiento_dai}* (${diasVencida}d atrás)`,
                    fechaSort: dVencimientoDai
                });
            }
        }
    }

    alertasEta.sort((a, b) => a.fechaSort - b.fechaSort);
    alertasVencimientos.sort((a, b) => a.fechaSort - b.fechaSort);

    return [...alertasVencimientos.map(i => i.texto), ...alertasEta.map(i => i.texto)];
}

async function procesarComandoLogistica(textoMensaje) {
    const textoLower = textoMensaje.toLowerCase();
    const palabrasEnMensaje = textoLower.split(/\s+/);

    if (textoLower.includes('ayuda') || textoLower.includes('comandos')) {
        return `🤖 *GUÍA DE COMANDOS - LOGISTATUS PRO*\n• "alertas activas"\n• "resumen exp"`;
    }
    if (textoLower.includes('alertas')) {
        const lista = await generarListaAlertasSistema();
        if (lista.length === 0) return "✅ No hay alertas activas en este momento.";
        return `🚨 *ALERTAS ACTIVAS (CON ART. 47 Y 177)* 🚨\n\n` + lista.join("\n\n");
    }
    return null;
}

async function enviarAlertaMulticanal(listaAlertas, tituloReporte, canalesForzados = null) {
    try {
        let enviarTg = canalesForzados ? canalesForzados.telegram : Boolean(TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID);
        let enviarWa = canalesForzados ? canalesForzados.whatsapp : Boolean(whatsappClientReady && sock);

        if (enviarTg && TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID) {
            let mensajesEnumerados = listaAlertas.map((item, index) => `${index + 1}️⃣ ${item}`);
            let mensajeConsolidado = `🚨 *LOGISTATUS PRO - ${tituloReporte}* 🚨\n\n` + mensajesEnumerados.join("\n\n");
            await enviarAlertaTelegramServidorPersonalizado(TELEGRAM_CHAT_ID, mensajeConsolidado);
        }

        if (enviarWa && whatsappClientReady && sock) {
            const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us";
            let mensajesEnumerados = listaAlertas.map((item, index) => `${index + 1}️⃣ ${item}`);
            let mensajeConsolidado = `🚨 *LOGISTATUS PRO - ${tituloReporte}* 🚨\n\n` + mensajesEnumerados.join("\n\n");
            await sock.sendMessage(GRUPO_AUTORIZADO_JID, { text: mensajeConsolidado });
        }
    } catch (err) {
        console.error("❌ Error en envío multicanal:", err);
    }
}

// ==========================================
// 6. RUTAS DE LA API Y CONFIGURACIÓN
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

        const expedientes = await obtenerExpedientesFirestore();
        const expedienteEncontrado = expedientes.find(e => String(e.num_expediente).trim() === String(numExpediente).trim());
        const datosAdjuntos = expedienteEncontrado ? [expedienteEncontrado] : [];

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

app.post('/api/alertas/forzar', async (req, res) => {
    try {
        const { canales } = req.body; 
        const listaFinalOrdenada = await generarListaAlertasSistema();
        if (listaFinalOrdenada.length === 0) {
            return res.json({ success: false, message: "No hay alertas activas." });
        }

        let canalesActivos = {
            whatsapp: Boolean(canales?.whatsapp && whatsappClientReady && sock),
            telegram: Boolean(canales?.telegram && TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID)
        };

        await enviarAlertaMulticanal(listaFinalOrdenada, "REPORTE FORZADO MANUAL (ORDENADO POR ETA)", canalesActivos);
        res.json({ success: true, message: "Alertas forzadas y enviadas con éxito.", canalesUtilizados: canalesActivos });
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

// ==========================================
// 7. FUNCIONES AUXILIARES Y TAREAS CRON
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

function formatearFechaDDMMYYYY(str) {
    const d = parseFechaLocalServidor(str);
    if (!d || isNaN(d.getTime())) return str || 'S/N';
    return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
}

// CRON: Alertas automáticas y correos con Excel a las 10:00 AM y 6:00 PM
cron.schedule('0 10,18 * * *', async () => {
    try {
        console.log("⏰ [CRON AUTOMÁTICO] Ejecutando reporte de alertas y correos con Excel...");
        
        const alertasGeneradas = await generarListaAlertasSistema();
        if (alertasGeneradas && alertasGeneradas.length > 0) {
            await enviarAlertaMulticanal(alertasGeneradas, "REPORTE AUTOMÁTICO PROGRAMADO");
        }

        const expedientes = await obtenerExpedientesFirestore();
        const clientes = await obtenerClientesFirestore();
        const clientesActivos = clientes.filter(c => c.enviarAutomatico !== false && c.correo);
        let enviadosCount = 0;

        for (const cliente of clientesActivos) {
            const expedientesCliente = expedientes.filter(e => 
                String(e.cliente).trim().toLowerCase() === String(cliente.razonSocial).trim().toLowerCase() && !e.despacho
            );

            if (expedientesCliente.length > 0) {
                const asunto = `Reporte Automático de Estatus - Logistatus Pro`;
                const mensaje = `Estimado/a ${cliente.razonSocial},\n\nAdjunto encontrará el reporte actualizado con el estatus de sus expedientes.\n\nEste es un mensaje automatizado generado por Logistatus Pro.`;

                await enviarCorreoConGoogleScriptDinamico(cliente.correo, asunto, mensaje, expedientesCliente);
                enviadosCount++;
            }
        }
        console.log(`✅ [CRON] Reportes automáticos enviados a ${enviadosCount} clientes.`);
    } catch (error) {
        console.error("❌ [CRON AUTOMÁTICO] Error:", error);
    }
}, { timezone: "America/Caracas" });

// ==========================================
// 8. INICIO DEL SERVIDOR
// ==========================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => { console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`); });dai