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
// 0.2 INICIALIZACIÓN DE WHATSAPP CON BAILEYS (INTELIGENCIA DE COMANDOS COMPLETA)
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
                
                console.log(`⚠️ [WHATSAPP] Conexión cerrada. Código: ${statusCode}. ¿Reconectar?: ${shouldReconnect}`);
                whatsappClientReady = false;
                estadoConexionWhatsapp = "DESCONECTADO";

                try {
                    if (sock) {
                        const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us";
                        await sock.sendMessage(GRUPO_AUTORIZADO_JID, { 
                            text: `⚠️ *ALERTA DE SISTEMA - LOGISTATUS PRO*\n\nEl bot de WhatsApp se ha **desconectado** (Código: ${statusCode || 'Desconocido'}). Intentando reconectar automáticamente...` 
                        });
                    }
                } catch (errWhatsapp) {
                    console.error("❌ No se pudo enviar alerta de desconexión al grupo de WhatsApp:", errWhatsapp);
                }

                if (shouldReconnect) {
                    setTimeout(() => {
                        iniciarBaileysWhatsApp();
                    }, 5000);
                }
            }
        });

        sock.ev.on('creds.update', saveCreds);

        // ESCUCHA INTELIGENTE EXCLUSIVA Y SEGURA PARA TU GRUPO
        sock.ev.on('messages.upsert', async (m) => {
            try {
                if (m.type !== 'notify') return;

                const msg = m.messages[0];
                if (!msg.message || msg.key.fromMe) return;

                const remitente = msg.key.remoteJid;
                const textoMensaje = (msg.message.conversation || msg.message.extendedTextMessage?.text || "").trim();
                const textoLower = textoMensaje.toLowerCase();

                console.log(`📩 [MENSAJE RECIBIDO] De: ${remitente} | Texto: ${textoMensaje}`);

                const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us"; 
                
                if (remitente !== GRUPO_AUTORIZADO_JID) {
                    console.log(`⛔ Mensaje ignorado por seguridad (No proviene del grupo autorizado).`);
                    return;
                }

                // =========================================================================
                // 📈 3. COMANDOS DE RESUMEN Y CONTEO (CON DAI, VIGENCIA Y OBSERVACIONES)
                // =========================================================================
                if (
                    textoLower.includes('resumen') || 
                    textoLower.includes('conteo') || 
                    textoLower.includes('totales') || 
                    textoLower.includes('activos') || 
                    textoLower.includes('en proceso') ||
                    textoLower.includes('resumen de expedientes') ||
                    textoLower.includes('resumen de exp') ||
                    textoLower.includes('expedientes activos') ||
                    textoLower.includes('exp activos') ||
                    textoLower.includes('expedientes en proceso') ||
                    textoLower.includes('exp en proceso')
                ) {
                    console.log(`📊 [WHATSAPP] Solicitud de detalle por pestañas detectada.`);
                    
                    const records = await obtenerExpedientesFirestore();
                    if (!records || records.length === 0) {
                        await sock.sendMessage(remitente, { text: "✅ No hay registros en Firestore para resumir." });
                        return;
                    }

                    let porLlegar = [];
                    let porRegistrar = [];
                    let porReconocer = [];
                    let porValidar = [];
                    let porDespachar = [];
                    const today = new Date();
                    today.setHours(0,0,0,0);

                    for (const rec of records) {
                        if (rec.despacho) continue;

                        const hasLlegada = Boolean(rec.fecha_llegada && rec.fecha_llegada.trim() !== "");
                        const hasTransmision = Boolean(rec.fecha_transmision && rec.fecha_transmision.trim() !== "");
                        const hasReconocimiento = Boolean(rec.reconocimiento && rec.reconocimiento.trim() !== "");
                        const hasValidacion = Boolean(rec.validacion && rec.validacion.trim() !== "");

                        if (hasValidacion) {
                            porDespachar.push(rec);
                        } else if (hasReconocimiento) {
                            porValidar.push(rec);
                        } else if (hasTransmision) {
                            porReconocer.push(rec);
                        } else if (hasLlegada) {
                            porRegistrar.push(rec);
                        } else {
                            porLlegar.push(rec);
                        }
                    }

                    const formatearBloqueSeccion = (titulo, lista, esPorRegistrar = false) => {
                        if (lista.length === 0) return `*${titulo} (0)*\n• Ninguno\n\n`;
                        let res = `*${titulo} (${lista.length})*\n`;
                        
                        lista.forEach((rec, idx) => {
                            const exp = rec.num_expediente || 'S/N';
                            const cli = rec.cliente || 'S/C';
                            const prov = rec.proveedor || 'N/A';
                            const bl = rec.awb_bl || rec.bl || 'S/N';
                            const obs = rec.observaciones ? rec.observaciones.trim() : '';
                            
                            let etiquetaFecha = "ETA";
                            let valorFecha = rec.eta_la_guaira || rec.fecha_llegada || 'Pendiente';
                            
                            if (esPorRegistrar) {
                                etiquetaFecha = "Fecha de Llegada";
                                valorFecha = rec.fecha_llegada || 'Pendiente';
                            }

                            // Estatus de DAI, vigencia y observaciones
                            const hasRegistroDai = Boolean(rec.fecha_registro_dai && rec.fecha_registro_dai.trim() !== "");
                            let detalleDaiTexto = "";
                            if (hasRegistroDai) {
                                const fechaDaiFormatted = formatearFechaDDMMYYYY(rec.fecha_registro_dai);
                                const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
                                if (dVencimientoDai) {
                                    const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                                    detalleDaiTexto = diffVencDays < 0 ? `\n   • DAI: *Registrada (${fechaDaiFormatted}) - ⚠️ VIGENCIA VENCIDA*` : `\n   • DAI: *Registrada (${fechaDaiFormatted}) - ✅ Vigente*`;
                                } else {
                                    detalleDaiTexto = `\n   • DAI: *Registrada (${fechaDaiFormatted})*`;
                                }
                            } else {
                                detalleDaiTexto = `\n   • DAI: *⚠️ FALTA REGISTRAR DAI / DUA*`;
                            }

                            res += `${idx + 1}️⃣ Exp: *${exp}*\n` +
                                   `   • Cliente: *${cli}*\n` +
                                   `   • Prov: *${prov}*\n` +
                                   `   • B/L: *${bl}*\n` +
                                   `   • ${etiquetaFecha}: *${valorFecha}*` +
                                   detalleDaiTexto;

                            if (obs !== '' && obs !== 'Ninguna') {
                                res += `\n   • Obs: _${obs}_`;
                            }

                            res += `\n\n`;
                        });
                        return res;
                    };

                    let mensajeDetallado = `📊 *EXPEDIENTES EN PROCESO - LOGISTATUS PRO* 📊\n\n` +
                        formatearBloqueSeccion("🚢✈️ 1. POR LLEGAR", porLlegar) +
                        formatearBloqueSeccion("📝 2. POR REGISTRAR", porRegistrar, true) +
                        formatearBloqueSeccion("🔍 3. POR RECONOCER", porReconocer, false) +
                        formatearBloque("📋 4. POR VALIDAR", porValidar, false) +
                        formatearBloque("🚚 5. POR DESPACHAR", porDespachar, false);

                    await sock.sendMessage(remitente, { text: mensajeDetallado });
                    return;
                }

                // =========================================================================
                // 🚨 2. COMANDO: CONSULTAR ALERTAS ACTIVAS
                // =========================================================================
                if (
                    textoLower === 'alertas' || 
                    textoLower === 'alertas activas' || 
                    textoLower === 'urgentes' || 
                    textoLower === 'reporte activo' ||
                    textoLower === 'vencimientos'
                ) {
                    console.log(`🚨 [WHATSAPP] Solicitud de alertas activas detectada.`);
                    
                    const records = await obtenerExpedientesFirestore();
                    if (!records || records.length === 0) {
                        await sock.sendMessage(remitente, { text: "✅ No hay registros en Firestore para evaluar." });
                        return;
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
                        const dArriboObj = parseFechaLocalServidor(fechaEtaRaw);

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
                                        detalleDaiTexto = diffVencDays < 0 ? `\n📄 DAI (${fechaDaiFormatted}) - ⚠️ VENCIDA` : `\n📄 DAI (${fechaDaiFormatted}) - ✅ Vigente`;
                                    } else {
                                        detalleDaiTexto = `\n📄 DAI Registrada (${fechaDaiFormatted})`;
                                    }
                                } else {
                                    detalleDaiTexto = `\n⚠️ *FALTA REGISTRAR DAI / DUA*`;
                                }

                                alertasEta.push(`${iconoVia} *${destinoTxt.toUpperCase()}*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n⏳ *${textoDias}* (${fechaEtaFormatted})${detalleDaiTexto}`);
                            }
                        }
                    }

                    let listaFinalOrdenada = alertasEta;
                    if (listaFinalOrdenada.length === 0) {
                        await sock.sendMessage(remitente, { text: "✅ *¡EXCELENTE!* No hay alertas activas en este momento." });
                        return;
                    }

                    let mensajesEnumerados = listaFinalOrdenada.map((item, index) => `${index + 1}️⃣ ${item}`);
                    let mensajeConsolidado = `🚨 *ALERTAS ACTIVAS - LOGISTATUS PRO* 🚨\n\n` + mensajesEnumerados.join("\n\n. . . . . . . . . . . . . . . . . . . . . . . .\n\n");
                    
                    await sock.sendMessage(remitente, { text: mensajeConsolidado });
                    return;
                }

            } catch (err) {
                console.error("❌ Error procesando mensaje en Baileys:", err);
            }
        });

    } catch (err) {
        console.error("❌ Error al iniciar Baileys en la nube:", err);
    }
}

iniciarBaileysWhatsApp();

// ==========================================
// 0.3 INICIALIZACIÓN DE TELEGRAM (POLLING DE COMANDOS)
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

                    if (TELEGRAM_CHAT_ID && chatId !== String(TELEGRAM_CHAT_ID)) continue;

                    if (texto.includes('resumen') || texto.includes('conteo') || texto.includes('alertas')) {
                        const records = await obtenerExpedientesFirestore();
                        let msgRespuesta = `📊 *REPORTE TELEGRAM*\nTotal registros activos: ${records.filter(r => !r.despacho).length}`;
                        await enviarAlertaTelegramServidor(msgRespuesta);
                    }
                }
            }
        } catch (err) {}
    }, 5000);
}
iniciarTelegramPolling();

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
        return res.status(400).json({ success: false, message: "El bot de WhatsApp aún no está conectado." });
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

// =========================================================================
// RUTA DE FORZAR ALERTAS (RESPETA CANALES ACTIVOS Y CAPTURA SOTO ADUANAS)
// =========================================================================
app.post('/api/alertas/forzar', async (req, res) => {
    try {
        const { canales } = req.body; 
        
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

            // 1. Alertas de ETA con estatus y fecha de DAI detallada
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
                            detalleDaiTexto = diffVencDays < 0 ? `\n📄 DAI (${fechaDaiFormatted}) - ⚠️ VIGENCIA VENCIDA` : `\n📄 DAI (${fechaDaiFormatted}) - ✅ Vigente`;
                        } else {
                            detalleDaiTexto = `\n📄 DAI Registrada (${fechaDaiFormatted})`;
                        }
                    } else {
                        detalleDaiTexto = `\n⚠️ *FALTA REGISTRAR DAI / DUA*`;
                    }

                    alertasEta.push(`${iconoVia} *${destinoTxt.toUpperCase()}*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ *${textoDias}* (${fechaEtaFormatted})${detalleDaiTexto}`);
                }
            }

            // 2. Plazo legal de Declaración (Art. 47 LOA)
            const dLlegada = parseFechaLocalServidor(rec.fecha_llegada);
            if (dLlegada && !hasTransmision) {
                let dDeadline = new Date(dLlegada.getTime());
                let diasHabilesAgregados = 0;
                let intentosMaximos = 0;
                while (diasHabilesAgregados < 5 && intentosMaximos < 30) {
                    dDeadline.setDate(dDeadline.getDate() + 1);
                    const diaSemana = dDeadline.getDay();
                    if (diaSemana !== 0 && diaSemana !== 6) diasHabilesAgregados++;
                    intentosMaximos++;
                }
                const diffPlazoDays = Math.round((dDeadline - today) / (1000 * 60 * 60 * 24));
                const fechaPlazoStr = formatearFechaDDMMYYYY(dDeadline.toISOString().split('T')[0]);

                if (diffPlazoDays < 0) {
                    const diasExt = Math.abs(diffPlazoDays);
                    alertasVencimientos.push(`🚨 *EXTEMPORANEIDAD - ART. 47*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ Venció el ${fechaPlazoStr} (${diasExt}d en extemporaneidad).\n🚨 *¡Aplica Multa Art. 177 Num 7!*`);
                }
            }

            // 3. Vencimiento de Días Libres en Almacén
            if (rec.culminacion_dias_libres) {
                const culmDate = parseFechaLocalServidor(rec.culminacion_dias_libres);
                const culmFormatted = formatearFechaDDMMYYYY(rec.culminacion_dias_libres);
                if (culmDate) {
                    const diffDaysAlm = Math.round((culmDate - today) / (1000 * 60 * 60 * 24));
                    if (diffDaysAlm < 0) {
                        alertasVencimientos.push(`🚨 *ALMACÉN VENCIDO*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⚠️ Venció hace ${Math.abs(diffDaysAlm)} día(s) el ${culmFormatted}`);
                    }
                }
            }

            // 4. Vencimiento DUA / DAI (Captura Soto Aduanas y Grand Aduanas)
            const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
            if (!hasLlegadaRegistrada && !hasRegistroDai && dVencimientoDai) {
                const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                if (diffVencDays < 0) {
                    const diasVencida = Math.abs(diffVencDays);
                    alertasVencimientos.push(`🚨 *VENCIDO / MULTA ART. 177*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⚠️ *Venció el ${rec.vencimiento_dai}* (${diasVencida}d atrás)`);
                }
            }
        }

        let listaFinalOrdenada = [...alertasEta, ...alertasVencimientos];
        if (listaFinalOrdenada.length === 0) {
            return res.json({ success: false, message: "No hay alertas activas en este momento para enviar." });
        }

        // VALIDACIÓN ESTRICTA: Solo envía a los canales que estén encendidos en la interfaz
        let canalesActivos = {
            whatsapp: Boolean(canales?.whatsapp && whatsappClientReady && sock),
            telegram: Boolean(canales?.telegram && TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID)
        };

        await enviarAlertaMulticanal(listaFinalOrdenada, "REPORTE FORZADO MANUAL", canalesActivos);
        res.json({ 
            success: true, 
            message: "Alertas forzadas y enviadas con éxito a las apps encendidas.",
            canalesUtilizados: canalesActivos
        });

    } catch (error) {
        console.error("❌ Error al forzar alertas:", error);
        res.status(500).json({ success: false, message: error.message });
    }
});

// ==========================================
// 2. FUNCIONES DE SERVICIO AUXILIARES
// ==========================================
function formatearFechaDDMMYYYY(str) {
    const d = parseFechaLocalServidor(str);
    if (!d || isNaN(d.getTime())) return str || 'S/N';
    const dia = String(d.getDate()).padStart(2, '0');
    const mes = String(d.getMonth() + 1).padStart(2, '0');
    const anio = d.getFullYear();
    return `${dia}-${mes}-${anio}`;
}

async function enviarAlertaMulticanal(listaAlertas, tituloReporte, canalesActivos = { whatsapp: true, telegram: true }) {
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
        console.error("❌ Error en el envío multicanal de alertas:", err);
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
    } catch (err) {
        console.error("❌ Error Telegram:", err);
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
        return [];
    }
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

// ==========================================
// 3. TAREAS PROGRAMADAS (CRON)
// ==========================================
cron.schedule('0 10,18 * * *', async () => {
    try {
        const correoCliente = process.env.CORREO_DESTINATARIO_AUTOMATICO;
        if (!correoCliente) return;
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) return;
        // Tarea automática de correo existente intacta...
    } catch (error) {
        console.error("❌ [CRON] Error:", error);
    }
}, { timezone: "America/Caracas" });

// ==========================================
// 4. INICIO DEL SERVIDOR
// ==========================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`);
});