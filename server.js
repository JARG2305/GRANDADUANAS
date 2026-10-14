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
// 0.2 INICIALIZACIÓN DE WHATSAPP CON BAILEYS (INTELIGENCIA DE COMANDOS)
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
                console.log('📱 [WHATSAPP BAILEYS] Nuevo código QR generado.');
                estadoConexionWhatsapp = "ESPERANDO_ESCANEO";
                try {
                    ultimoQrGenerado = await qrcode.toDataURL(qr);
                } catch (err) {
                    console.error("❌ Error al convertir QR de Baileys a DataURL:", err);
                }
            }

            if (connection === 'open') {
                console.log('✅ [WHATSAPP BAILEYS] ¡Conectado y listo en la nube!');
                whatsappClientReady = true;
                estadoConexionWhatsapp = "CONECTADO";
                ultimoQrGenerado = "";
            }

            if (connection === 'close') {
                const shouldReconnect = lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut;
                console.log(`⚠️ [WHATSAPP BAILEYS] Conexión cerrada. Reconectando: ${shouldReconnect}`);
                whatsappClientReady = false;
                estadoConexionWhatsapp = "DESCONECTADO";
                if (shouldReconnect) {
                    iniciarBaileysWhatsApp();
                }
            }
        });

        sock.ev.on('creds.update', saveCreds);

        // ESCUCHA INTELIGENTE PURA POR COINCIDENCIA DE EVENTO + EXPEDIENTE / B/L
        sock.ev.on('messages.upsert', async (m) => {
            try {
                const msg = m.messages[0];
                if (!msg.message || msg.key.fromMe) return;

                const textoMensaje = (msg.message.conversation || msg.message.extendedTextMessage?.text || "").trim();
                const remitente = msg.key.remoteJid;
                const textoLower = textoMensaje.toLowerCase();

                // 🔍 LÍNEA DE DEPURACIÓN PARA VER EL CHAT O GRUPO EN LOS LOGS
                console.log(`📩 [MENSAJE RECIBIDO] De: ${remitente} | Texto: ${textoMensaje}`);

                // 1. Identificar el campo operativo según la palabra clave o evento en el mensaje
                let campoAModificar = "";
                if (textoLower.includes('eta')) {
                    campoAModificar = "eta_la_guaira";
                } else if (textoLower.includes('llegada')) {
                    campoAModificar = "fecha_llegada";
                } else if (textoLower.includes('transmision') || textoLower.includes('transmitido')) {
                    campoAModificar = "fecha_transmision";
                } else if (textoLower.includes('dai')) {
                    campoAModificar = "fecha_registro_dai";
                } else if (textoLower.includes('acta')) {
                    campoAModificar = "recibida_acta_recepcion";
                }

                // Si el mensaje no contiene ningún evento conocido, lo ignoramos para que no interfiera con chats normales
                if (!campoAModificar) return;

                // 2. Extraer fecha (ej: 08/11/2026, 08-11-2026) o usar la de hoy por defecto en Venezuela
                const matchFecha = textoMensaje.match(/\b(\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})\b/);
                let fechaDetectada = matchFecha ? matchFecha[0].replace(/-/g, '/') : new Date().toLocaleDateString("es-VE", { timeZone: "America/Caracas" });

                // 3. Consultar Firestore para obtener todos los expedientes y buscar la coincidencia de texto
                const projectId = process.env.FIREBASE_PROJECT_ID;
                const urlQuery = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/expedientes`;
                const response = await fetch(urlQuery);
                const result = await response.json();

                if (!result.documents || result.documents.length === 0) {
                    await sock.sendMessage(remitente, { text: "❌ No se encontraron expedientes en la base de datos." });
                    return;
                }

                // Convertir todos los documentos de Firestore a un formato plano fácil de buscar
                let listaExpedientes = result.documents.map(doc => {
                    let fields = doc.fields || {};
                    return {
                        id: doc.name.split('/').pop(),
                        fields: fields,
                        num_expediente: (fields.num_expediente?.stringValue || "").toLowerCase(),
                        bl: (fields.awb_bl?.stringValue || fields.bl?.stringValue || fields.numero_bl?.stringValue || "").toLowerCase()
                    };
                });

                // 4. Buscar cuál es el expediente o B/L mencionado dentro del texto del mensaje
                let expedienteEncontrado = null;
                
                for (let item of listaExpedientes) {
                    if ((item.num_expediente && item.num_expediente.length > 1 && textoLower.includes(item.num_expediente)) ||
                        (item.bl && item.bl.length > 2 && textoLower.includes(item.bl))) {
                        expedienteEncontrado = item;
                        break;
                    }
                }

                if (!expedienteEncontrado) {
                    await sock.sendMessage(remitente, { text: "⚠️ No pude identificar ningún número de expediente o B/L válido en tu mensaje para actualizar el evento." });
                    return;
                }

                let numExpReal = expedienteEncontrado.fields.num_expediente?.stringValue || 'S/N';
                let blReal = expedienteEncontrado.fields.awb_bl?.stringValue || expedienteEncontrado.fields.bl?.stringValue || 'S/N';

                // 5. Actualizar en Firestore mediante petición PATCH preservando el resto de los campos
                const updateUrl = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/expedientes/${expedienteEncontrado.id}?updateMask.fieldPaths=${campoAModificar}`;
                
                const patchBody = {
                    fields: {
                        ...expedienteEncontrado.fields,
                        [campoAModificar]: { stringValue: fechaDetectada }
                    }
                };

                const patchRes = await fetch(updateUrl, {
                    method: 'PATCH',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(patchBody)
                });

                if (patchRes.ok) {
                    await sock.sendMessage(remitente, { text: `✅ *¡ACTUALIZACIÓN AUTOMÁTICA EXITOSA!* 🚀\n\n📂 Exp: *${numExpReal}*\n📦 B/L / AWB: *${blReal}*\n⚙️ Evento/Campo: *${campoAModificar}*\n📅 Fecha Registrada: *${fechaDetectada}*` });
                } else {
                    const errData = await patchRes.json();
                    await sock.sendMessage(remitente, { text: `❌ Error al actualizar en Firestore: ${errData.error?.message || 'Desconocido'}` });
                }

            } catch (err) {
                console.error("❌ Error procesando mensaje inteligente por coincidencia en Baileys:", err);
            }
        });

    } catch (err) {
        console.error("❌ Error al iniciar Baileys en la nube:", err);
    }
}
// Iniciar el cliente de WhatsApp al arrancar el servidor
iniciarBaileysWhatsApp();

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

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/health', (req, res) => {
    res.status(200).send('LOGISTATUS PRO - Servidor Activo OK');
});

// ==========================================
// 1.1 ENDPOINTS DE WHATSAPP PARA LA WEB
// ==========================================
app.get('/api/whatsapp/estado', (req, res) => {
    res.json({
        success: true,
        estado: estadoConexionWhatsapp,
        ready: whatsappClientReady,
        qrCodeDataUrl: ultimoQrGenerado
    });
});

app.post('/api/whatsapp/enviar', async (req, res) => {
    const { telefono, mensaje } = req.body;
    if (!whatsappClientReady || !sock) {
        return res.status(400).json({ success: false, message: "El bot de WhatsApp aún no está conectado. Escanee el código QR en el sistema." });
    }
    if (!telefono || !mensaje) {
        return res.status(400).json({ success: false, message: "Faltan el número de teléfono o el mensaje." });
    }

    try {
        let numeroLimpios = telefono.replace(/\D/g, '');
        const jid = `${numeroLimpios}@s.whatsapp.net`;

        await sock.sendMessage(jid, { text: mensaje });
        res.json({ success: true, message: "Mensaje de WhatsApp enviado con éxito." });
    } catch (error) {
        console.error("❌ Error al enviar WhatsApp con Baileys:", error);
        res.status(500).json({ success: false, message: error.message });
    }
});

app.all('/api/calcular-alertas', async (req, res) => {
    try {
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) {
            return res.json({ success: true, count: 0, message: "No hay registros en Firestore." });
        }

        let alertasEta = [];
        let alertasDaiPendientes = [];
        let alertasVencimientos = [];
        
        const today = new Date();
        today.setHours(0,0,0,0);

        for (const rec of records) {
            if (rec.despacho) continue;

            const expName = rec.num_expediente || 'S/N';
            const clientName = rec.cliente || 'S/C';
            const blInfo = rec.awb_bl || rec.bl || rec.numero_bl || 'S/N'; 
            const fechaEta = rec.fecha_llegada || rec.eta_la_guaira || 'S/N';
            const modoBruto = (rec.modo_transporte || "").toLowerCase().trim();
            const esBarco = !modoBruto.includes("aéreo") && !modoBruto.includes("aereo");
            const iconoVia = esBarco ? "🚢" : "✈️";
            const destinoTxt = esBarco ? "ETA La Guaira" : "ETA Maiquetía";

            const fechaLlegadaStr = rec.fecha_llegada;
            const hasLlegadaRegistrada = Boolean(fechaLlegadaStr && fechaLlegadaStr.trim() !== "");
            const hasRegistroDai = Boolean(rec.fecha_registro_dai && rec.fecha_registro_dai.trim() !== "");
            const hasTransmision = Boolean(rec.fecha_transmision && rec.fecha_transmision.trim() !== "");

            const fechaArriboObjetivo = rec.fecha_llegada || rec.eta_la_guaira;
            const dArriboObj = parseFechaLocalServidor(fechaArriboObjetivo);

            if (dArriboObj && !hasLlegadaRegistrada) {
                const diffLlegadaDays = Math.round((dArriboObj - today) / (1000 * 60 * 60 * 24));
                if (diffLlegadaDays >= 0 && diffLlegadaDays <= 5) {
                    const textoDias = diffLlegadaDays === 0 ? "¡LLEGA HOY!" : `Faltan ${diffLlegadaDays}d`;
                    let tarjeta = `${iconoVia} *${destinoTxt.toUpperCase()}*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ *${textoDias}* (${fechaArriboObjetivo})`;
                    alertasEta.push(tarjeta);
                }
            }

            if (dArriboObj && !hasLlegadaRegistrada && !hasRegistroDai) {
                const diffEtaDaiDays = Math.round((dArriboObj - today) / (1000 * 60 * 60 * 24));
                if (diffEtaDaiDays >= 0 && diffEtaDaiDays <= 5) {
                    const textoEtaDai = diffEtaDaiDays === 0 ? "¡LLEGA HOY!" : `Faltan ${diffEtaDaiDays}d para el ETA`;
                    let tarjeta = `⚠️ *DAI POR REGISTRAR (ETA PRÓXIMO)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoEtaDai} (${fechaArriboObjetivo})\n⚠️ *La fecha de registro DAI se encuentra en blanco.*`;
                    alertasDaiPendientes.push(tarjeta);
                }
            }

            const dLlegada = parseFechaLocalServidor(rec.fecha_llegada);
            
            if (dLlegada && !hasTransmision) {
                let dDeadline = new Date(dLlegada.getTime());
                let diasHabilesAgregados = 0;
                let intentosMaximos = 0;
                while (diasHabilesAgregados < 5 && intentosMaximos < 30) {
                    dDeadline.setDate(dDeadline.getDate() + 1);
                    const diaSemana = dDeadline.getDay();
                    if (diaSemana !== 0 && diaSemana !== 6) {
                        diasHabilesAgregados++;
                    }
                    intentosMaximos++;
                }
                const diffPlazoDays = Math.round((dDeadline - today) / (1000 * 60 * 60 * 24));
                const fechaPlazoStr = dDeadline.toISOString().split('T')[0];

                if (diffPlazoDays >= 0 && diffPlazoDays <= 2) {
                    const textoPlazo = diffPlazoDays === 0 ? "¡VENCE HOY!" : `Quedan ${diffPlazoDays} días para transmitir`;
                    let tarjeta = `⚠️ *PLAZO DE DECLARACIÓN PRÓXIMO A VENCER (Art. 47 LOA)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoPlazo} - *Riesgo Multa Art. 177 Num. 7 de la L.O.A.*`;
                    alertasVencimientos.push(tarjeta);
                } else if (diffPlazoDays < 0) {
                    const diasExt = Math.abs(diffPlazoDays);
                    let tarjeta = `🚨 *EXTEMPORANEIDAD - ART. 47 / 177 Num. 7 de la L.O.A.*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ Plazo vencido el ${fechaPlazoStr} (Vencida hace ${diasExt} día(s)).\n🚨 *¡Aplica Multa Art. 177 N.º 7 de la L.O.A.! (Persiste hasta transmitir)*`;
                    alertasVencimientos.push(tarjeta);
                }
            }

            if (rec.culminacion_dias_libres) {
                const culmDate = parseFechaLocalServidor(rec.culminacion_dias_libres);
                if (culmDate) {
                    const diffDaysAlm = Math.round((culmDate - today) / (1000 * 60 * 60 * 24));
                    if (diffDaysAlm >= 0 && diffDaysAlm <= 2) {
                        const textoAlm = diffDaysAlm === 0 ? "VENCE HOY" : `Faltan ${diffDaysAlm}d`;
                        let tarjeta = `⚠️ *ALMACÉN PRÓXIMO A VENCER*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoAlm} (${rec.culminacion_dias_libres})`;
                        alertasVencimientos.push(tarjeta);
                    } else if (diffDaysAlm < 0) {
                        let tarjeta = `🚨 *ALMACÉN VENCIDO*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⚠️ Venció hace ${Math.abs(diffDaysAlm)} día(s) el ${rec.culminacion_dias_libres}`;
                        alertasVencimientos.push(tarjeta);
                    }
                }
            }

            const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
            if (!hasLlegadaRegistrada && !hasRegistroDai && dVencimientoDai) {
                const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                
                if (diffVencDays < 0) {
                    const diasVencida = Math.abs(diffVencDays);
                    let tarjeta = `🚨 *DAI VENCIDA (SIN REGISTRAR)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⚠️ Venció el ${rec.vencimiento_dai} (${diasVencida} día(s) atrás) sin registrar DAI.\n🚨 *¡Aplica Multa Art. 177 Num. 9 de la L.O.A.! (Persiste hasta registrar)*`;
                    alertasVencimientos.push(tarjeta);
                } else if (diffVencDays >= 0 && diffVencDays <= 4) {
                    const textoCountdown = diffVencDays === 0 ? "¡VENCE HOY!" : `Faltan ${diffVencDays}d`;
                    let tarjeta = `⚠️ *DAI PRÓXIMA A VENCER*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoCountdown} (${rec.vencimiento_dai})\n🚢 ETA: *${fechaEta}* - *Riesgo Multa Art 177 Num 9 de la L.O.A*`;
                    alertasVencimientos.push(tarjeta);
                }
            }
        }

        let listaFinalOrdenada = [...alertasEta, ...alertasDaiPendientes, ...alertasVencimientos];

        if (listaFinalOrdenada.length === 0) {
            return res.json({ success: true, count: 0, message: "No hay alertas críticas pendientes." });
        }

        let mensajesEnumerados = listaFinalOrdenada.map((item, index) => `${index + 1}️⃣ ${item}`);

        let mensajeConsolidado = "🚨 *LOGISTATUS PRO - REPORTE CONSOLIDADO* 🚨\n\n" + mensajesEnumerados.join("\n\n. . . . . . . . . . . . . . . . . . . . . . . .\n\n");
        await enviarAlertaTelegramServidor(mensajeConsolidado);

        res.json({ success: true, count: listaFinalOrdenada.length, message: `¡Alertas consolidadas y enviadas con éxito! (${listaFinalOrdenada.length} enviadas)` });
    } catch (error) {
        console.error("❌ Error en /api/calcular-alertas:", error);
        res.status(500).json({ success: false, message: error.message });
    }
});

// ==========================================
// 2. FUNCIONES DE SERVICIO (CORREO, TELEGRAM Y FIRESTORE)
// ==========================================
async function enviarCorreoSistema(opcionesMail) {
    try {
        const apiKey = process.env.RESEND_API_KEY; 
        const remitente = process.env.RESEND_REMITE || "onboarding@resend.dev";

        if (!apiKey || apiKey.includes("TU_API_KEY")) {
            throw new Error("Falta configurar la RESEND_API_KEY en las variables de entorno de Render para este cliente.");
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
    } catch (err) {
        console.error("❌ Error al consultar Firestore desde el servidor:", err);
        return [];
    }
}

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
cron.schedule('0 10,18 * * *', async () => {
    console.log("⏰ [CRON] Ejecutando generación y envío automático...");
    try {
        const correoCliente = process.env.CORREO_DESTINATARIO_AUTOMATICO;
        if (!correoCliente) return;
        const destinatariosAutomaticos = [correoCliente];
        
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) return;

        const excelHeaders = [
            "MODO VÍA", "N.º EXPEDIENTE", "CLIENTE", "PROVEEDOR", "LÍNEA", "PAÍS ORIGEN", "BUQUE / VUELO ORIGEN", "AWB / BL", "N.º CONTENEDOR(ES)", "PESO BL", "CONTENIDO SEGÚN BL", "ETD ORIGEN", "PUERTO TRANSBORDO", "ETA TRANSBORDO", "ETD TRANSBORDO", "BUQUE / VUELO A VE", "ETA LA GUAIRA", "FECHA REGISTRO DAI", "VENCIMIENTO DAI", "FECHA DE LLEGADA", "RECIBIDA ACTA RECEPCIÓN", "FECHA ABANDONO LEGAL", "PERMISOLOGÍA", "FECHA RECIBIDO PERMISOLOGÍA", "PREVALORACIÓN ENVIADA", "DOC. VALORADO EN SISTEMA", "FACTURA RECIBIDA", "MONTO FLETE", "RECIBIDO DOC. TRANSPORTE", "FECHA TRANSMISIÓN", "CANAL", "FUNCIONARIO", "RECONOCIMIENTO", "VALIDACIÓN", "DESPACHO", "ALMACÉN", "RECIBIDA ACTA (ALMACÉN)", "DÍAS LIBRES ALMACÉN", "INICIO DÍAS LIBRES ALMACÉN", "CULMINACIÓN DÍAS LIBRES ALMACÉN", "NAVIERA", "DÍAS LIBRES NAVIERA", "INICIO DÍAS LIBRES NAVIERA", "CULMINACIÓN DÍAS LIBRES NAVIERA", "OBSERVACIONES"
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

        const blackBorder = {
            top: { style: "thin", color: { rgb: "FFFFFF" } },
            bottom: { style: "thin", color: { rgb: "FFFFFF" } },
            left: { style: "thin", color: { rgb: "FFFFFF" } },
            right: { style: "thin", color: { rgb: "FFFFFF" } }
        };

        const anchosEspecificosXLSX = [
            10, 11, 31, 21, 15, 15, 21, 21, 18, 12, 18, 13, 14, 14, 14, 15, 14, 
            14.5, 14.5, 14.5, 16, 16, 12, 14.5, 14.5, 14.5, 16, 12, 12, 12, 
            14, 14, 14, 17, 15, 15, 17, 14, 14, 14, 15, 17, 14.5, 14.5, 14.5, 14.5, 20
        ];

        const buildRowsServidor = (listaRecs) => {
            return listaRecs.map(rec => fieldsList.map(f => rec[f] !== undefined && rec[f] !== null ? rec[f] : ""));
        };

        const ordenarPorETA = (lista) => {
            return lista.sort((a, b) => {
                let fA = parseFechaLocalServidor(a.eta_la_guaira || a.fecha_llegada) || new Date('9999-12-31');
                let fB = parseFechaLocalServidor(b.eta_la_guaira || b.fecha_llegada) || new Date('9999-12-31');
                return fA - fB;
            });
        };

        const porLlegar = ordenarPorETA(records.filter(r => !r.despacho && !r.validacion && !r.reconocimiento));
        const tramites = ordenarPorETA(records.filter(r => !r.despacho && (r.validacion || r.reconocimiento || r.fecha_transmision)));
        const despachados = ordenarPorETA(records.filter(r => r.despacho));

        let masterRows = [];
        if (porLlegar.length > 0) {
            masterRows.push(["▶ 1. EXPEDIENTES POR LLEGAR"]);
            masterRows.push(excelHeaders);
            masterRows.push(...buildRowsServidor(porLlegar));
            masterRows.push([]);
        }
        if (tramites.length > 0) {
            masterRows.push(["▶ 2. TRÁMITES ADUANALES Y EN PROCESO"]);
            masterRows.push(excelHeaders);
            masterRows.push(...buildRowsServidor(tramites));
            masterRows.push([]);
        }
        if (despachados.length > 0) {
            masterRows.push(["▶ 3. EXPEDIENTES DESPACHADOS"]);
            masterRows.push(excelHeaders);
            masterRows.push(...buildRowsServidor(despachados));
        }

        const wb = XLSX.utils.book_new();
        const wsMaster = XLSX.utils.aoa_to_sheet(masterRows);
        wsMaster['!cols'] = anchosEspecificosXLSX.map(w => ({ wch: w }));

        let currentSection = 1;
        const range = XLSX.utils.decode_range(wsMaster['!ref']);
        for (let R = range.s.r; R <= range.e.r; ++R) {
            const cellA = wsMaster[XLSX.utils.encode_cell({ r: R, c: 0 })];
            if (cellA && cellA.v && String(cellA.v).includes("▶")) {
                if (String(cellA.v).includes("TRÁMITES")) currentSection = 2;
                if (String(cellA.v).includes("DESPACHADOS")) currentSection = 3;
                cellA.s = { font: { name: "Segoe UI", sz: 11, bold: true, color: { rgb: "1E3A8A" } } };
                continue;
            }

            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                if (!wsMaster[cellAddress]) continue;

                let cellStyle = { border: blackBorder, alignment: { vertical: "center", horizontal: "center", wrapText: true }, font: { name: "Segoe UI", sz: 10 } };
                const headerCheck = wsMaster[XLSX.utils.encode_cell({ r: R, c: 0 })];
                if (headerCheck && excelHeaders.includes(headerCheck.v)) {
                    cellStyle.fill = { fgColor: { rgb: "1E3A8A" } };
                    cellStyle.font = { name: "Segoe UI", sz: 10, bold: true, color: { rgb: "FFFFFF" } };
                } else {
                    cellStyle.font = { name: "Segoe UI", sz: 10, color: { rgb: "000000" }, bold: true };
                    cellStyle.fill = { fgColor: { rgb: currentSection === 3 ? "FEE2E2" : currentSection === 2 ? "D9E1F2" : "F8FAFC" } };
                }
                wsMaster[cellAddress].s = cellStyle;
            }
        }

        XLSX.utils.book_append_sheet(wb, wsMaster, "Reporte Consolidado");
        const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
        const excelBase64String = Buffer.from(excelBuffer).toString('base64');

        await enviarCorreoSistema({
            to: [correoCliente],
            subject: "📊 Reporte Maestro Automático - LOGISTATUS PRO",
            text: "Adjunto el reporte automático consolidado.",
            attachments: [{ filename: 'Reporte_Automatico.xlsx', content: excelBase64String, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }]
        });
    } catch (error) {
        console.error("❌ [CRON] Error al enviar reporte automático:", error);
    }
}, { timezone: "America/Caracas" });

// Evaluación de alertas operativas
cron.schedule('0 9,12,15,18,19,20 * * *', async () => {
    try {
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) return;

        let alertasEta = [], alertasDaiPendientes = [], alertasVencimientos = [];
        const today = new Date();
        today.setHours(0,0,0,0);

        for (const rec of records) {
            if (rec.despacho) continue;

            const expName = rec.num_expediente || 'S/N';
            const clientName = rec.cliente || 'S/C';
            const blInfo = rec.awb_bl || rec.bl || rec.numero_bl || 'S/N'; 
            const fechaEta = rec.fecha_llegada || rec.eta_la_guaira || 'S/N';
            const modoBruto = (rec.modo_transporte || "").toLowerCase().trim();
            const esBarco = !modoBruto.includes("aéreo") && !modoBruto.includes("aereo");
            const iconoVia = esBarco ? "🚢" : "✈️";
            const destinoTxt = esBarco ? "ETA La Guaira" : "ETA Maiquetía";

            const hasLlegadaRegistrada = Boolean(rec.fecha_llegada && rec.fecha_llegada.trim() !== "");
            const hasRegistroDai = Boolean(rec.fecha_registro_dai && rec.fecha_registro_dai.trim() !== "");
            const dArriboObj = parseFechaLocalServidor(fechaEta);

            if (dArriboObj && !hasLlegadaRegistrada) {
                const diffLlegadaDays = Math.round((dArriboObj - today) / (1000 * 60 * 60 * 24));
                if (diffLlegadaDays >= 0 && diffLlegadaDays <= 5) {
                    const textoDias = diffLlegadaDays === 0 ? "¡LLEGA HOY!" : `Faltan ${diffLlegadaDays}d`;
                    alertasEta.push(`${iconoVia} *${destinoTxt.toUpperCase()}*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ *${textoDias}* (${fechaEta})`);
                }
            }

            if (dArriboObj && !hasLlegadaRegistrada && !hasRegistroDai) {
                const diffEtaDaiDays = Math.round((dArriboObj - today) / (1000 * 60 * 60 * 24));
                if (diffEtaDaiDays >= 0 && diffEtaDaiDays <= 5) {
                    const textoEtaDai = diffEtaDaiDays === 0 ? "¡LLEGA HOY!" : `Faltan ${diffEtaDaiDays}d para el ETA`;
                    alertasDaiPendientes.push(`⚠️ *DAI POR REGISTRAR*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoEtaDai}\n⚠️ *La fecha DAI está en blanco.*`);
                }
            }
        }

        let listaFinalOrdenada = [...alertasEta, ...alertasDaiPendientes, ...alertasVencimientos];
        if (listaFinalOrdenada.length > 0) {
            let mensajesEnumerados = listaFinalOrdenada.map((item, index) => `${index + 1}️⃣ ${item}`);
            let mensajeConsolidado = "🚨 *LOGISTATUS PRO - REPORTE AUTOMÁTICO* 🚨\n\n" + mensajesEnumerados.join("\n\n. . . . . . . . . . . . . . . . . . . . . . . .\n\n");
            await enviarAlertaTelegramServidor(mensajeConsolidado);
        }
    } catch (error) {
        console.error("❌ Error en cron de alertas:", error);
    }
}, { timezone: "America/Caracas" });

// ==========================================
// 5. INICIO DEL SERVIDOR
// ==========================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`);
});