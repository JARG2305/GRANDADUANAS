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
// 0.2 INICIALIZACIÓN DE WHATSAPP CON BAILEYS
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
// 0.3 INICIALIZACIÓN DE TELEGRAM (POLLING + CRON MULTICANAL)
// ==========================================
async function iniciarTelegramPolling() {
    if (!TELEGRAM_BOT_TOKEN) return;
    let offset = 0;
    
    console.log("🤖 [TELEGRAM] Iniciando sondeo (Polling) de comandos...");

    // Alertas Multicanal: Todos los días a las 9, 12, 15, 18, 19, 20 horas
    cron.schedule('0 9,12,15,18,19,20 * * *', async () => {
        try {
            console.log("⏰ [CRON MULTICANAL] Ejecutando reporte automático de alertas...");
            const alertasGeneradas = await generarListaAlertasSistema();
            if (alertasGeneradas && alertasGeneradas.length > 0) {
                await enviarAlertaMulticanal(alertasGeneradas, "REPORTE AUTOMÁTICO PROGRAMADO");
            }
        } catch (error) {
            console.error("❌ [CRON MULTICANAL] Error en reporte automático:", error);
        }
    }, { timezone: "America/Caracas" });

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
// 0.4 MOTOR CENTRAL DE COMANDOS Y ALERTAS (VALIDACIÓN DUAL EXP / BL)
// ==========================================
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

                alertasEta.push({
                    texto: `${iconoVia} *${destinoTxt.toUpperCase()}*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n⏳ *${textoDias}* (${fechaEtaFormatted})${detalleDaiTexto}`,
                    fechaSort: dArriboObj
                });
            }
        }

        const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
        if (dVencimientoDai) {
            const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
            if (diffVencDays < 0) {
                const diasVencida = Math.abs(diffVencDays);
                alertasVencimientos.push({
                    texto: `🚨 *VENCIDO / MULTA ART. 177 (DAI)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⚠️ *La DAI venció el ${rec.vencimiento_dai}* (${diasVencida}d atrás)`,
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

    // 1. GUÍA DE COMANDOS
    const palabrasAyuda = ['comandos', 'ayuda', 'menu', 'menú', 'sistema'];
    const coincidenciasAyuda = palabrasEnMensaje.filter(p => palabrasAyuda.includes(p)).length;

    if (coincidenciasAyuda >= 2 || textoLower === 'ayuda' || textoLower === 'comandos') {
        return `🤖 *GUÍA DE COMANDOS - LOGISTATUS PRO* 🤖\n\n` +
               `🚨 *1. Alertas Activas*\n` +
               `• *"alertas activas"*, *"alertas exp"*, *"alerta sistema"*\n\n` +
               `📊 *2. Resumen y Conteo*\n` +
               `• *"resumen exp"*, *"resumen expedientes"*, *"expedientes activos"*\n\n` +
               `✏️ *3. Actualización de Fechas (usa Exp o BL - Si omites fecha toma el día de hoy)*\n` +
               `• *ETA:* "Actualizar ETA exp 26-030"\n` +
               `• *Llegada:* "Actualiza llegada bl 12345 al 26/10/2026"\n` +
               `• *Transmitido:* "Transmitido bl 12345"\n` +
               `• *Reconocido:* "Reconocido exp 26-030"\n` +
               `• *Validado:* "Validado bl 12345"\n` +
               `• *Despachado:* "Despachado exp 26-030"\n` +
               `• *Retiro de acta:* "Retirada acta bl 12345"\n\n` +
               `✅ *4. Indicadores a SI*\n` +
               `• *"Retirado bl bl 12345"* | *"Prevaloración enviada exp 26-030"* | *"Valorado bl 12345"* | *"Factura recibida bl 12345"*\n\n` +
               `🟢 *5. Canal (Verde / Amarillo / Rojo)*\n` +
               `• *"Canal exp 26-030 verde"* | *"Canal bl 12345 rojo"*\n\n` +
               `🗑️ *6. Borrar Campo*\n` +
               `• *"Borrar ETA exp 26-030"* | *"Eliminar canal bl 12345"*\n\n` +
               `⚙️ *Ayuda:* "menu sistema"`;
    }

    // 2. CONSULTAR ALERTAS ACTIVAS
    const palabrasAlertas = ['alertas', 'alerta', 'activas', 'exp', 'expedientes', 'sistema', 'urgentes', 'vencimientos'];
    const coincidenciasAlertas = palabrasEnMensaje.filter(p => palabrasAlertas.includes(p)).length;

    if (coincidenciasAlertas >= 2) {
        const listaFinalOrdenada = await generarListaAlertasSistema();
        if (listaFinalOrdenada.length === 0) return "✅ *¡EXCELENTE!* No hay alertas activas.";
        let mensajesEnumerados = listaFinalOrdenada.map((item, index) => `${index + 1}️⃣ ${item}`);
        return `🚨 *ALERTAS ACTIVAS (ORDENADAS POR ETA/VENCIMIENTO)* 🚨\n\n` + mensajesEnumerados.join("\n\n");
    }

    // 3. CONSULTAR RESUMEN / CONTEO
    const palabrasResumen = ['resumen', 'exp', 'expedientes', 'activos', 'conteo', 'totales', 'proceso'];
    const coincidenciasResumen = palabrasEnMensaje.filter(p => palabrasResumen.includes(p)).length;

    if (coincidenciasResumen >= 2) {
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) return "✅ No hay registros en Firestore.";

        let porLlegar = [], porRegistrar = [], porReconocer = [], porValidar = [], porDespachar = [];
        const today = new Date();
        today.setHours(0,0,0,0);

        for (const rec of records) {
            if (rec.despacho) continue;
            const hasLlegada = Boolean(rec.fecha_llegada && rec.fecha_llegada.trim() !== "");
            const hasTransmision = Boolean(rec.fecha_transmision && rec.fecha_transmision.trim() !== "");
            const hasReconocimiento = Boolean(rec.reconocimiento && rec.reconocimiento.trim() !== "");
            const hasValidacion = Boolean(rec.validacion && rec.validacion.trim() !== "");

            if (hasValidacion) porDespachar.push(rec);
            else if (hasReconocimiento) porValidar.push(rec);
            else if (hasTransmision) porReconocer.push(rec);
            else if (hasLlegada) porRegistrar.push(rec);
            else porLlegar.push(rec);
        }

        const ordenarPorETA = (lista) => {
            return lista.sort((a, b) => {
                let fA = parseFechaLocalServidor(a.eta_la_guaira || a.fecha_llegada) || new Date('9999-12-31');
                let fB = parseFechaLocalServidor(b.eta_la_guaira || b.fecha_llegada) || new Date('9999-12-31');
                return fA - fB;
            });
        };

        porLlegar = ordenarPorETA(porLlegar);
        porRegistrar = ordenarPorETA(porRegistrar);
        porReconocer = ordenarPorETA(porReconocer);
        porValidar = ordenarPorETA(porValidar);
        porDespachar = ordenarPorETA(porDespachar);

        const formatearBloqueSeccion = (titulo, lista, esPorRegistrar = false, mostrarDai = false) => {
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

                res += `${idx + 1}️⃣ Exp: *${exp}*\n` +
                       `   • Cliente: *${cli}*\n` +
                       `   • Prov: *${prov}*\n` +
                       `   • B/L: *${bl}*\n` +
                       `   • ${etiquetaFecha}: *${valorFecha}*`;

                if (mostrarDai) {
                    const hasRegistroDai = Boolean(rec.fecha_registro_dai && rec.fecha_registro_dai.trim() !== "");
                    if (hasRegistroDai) {
                        const fechaDaiFormatted = formatearFechaDDMMYYYY(rec.fecha_registro_dai);
                        const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
                        if (dVencimientoDai) {
                            const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                            res += diffVencDays < 0 ? `\n   • DAI: *Registrada (${fechaDaiFormatted}) - ⚠️ VENCIDA*` : `\n   • DAI: *Registrada (${fechaDaiFormatted}) - ✅ Vigente*`;
                        } else {
                            res += `\n   • DAI: *Registrada (${fechaDaiFormatted})*`;
                        }
                    } else {
                        res += `\n   • DAI: *⚠️ FALTA REGISTRAR DAI / DUA*`;
                    }
                }

                if (obs !== '' && obs !== 'Ninguna') res += `\n   • Obs: _${obs}_`;
                res += `\n\n`;
            });
            return res;
        };

        return `📊 *EXPEDIENTES EN PROCESO (ORDENADOS POR ETA)* 📊\n\n` +
               formatearBloqueSeccion("🚢✈️ 1. POR LLEGAR", porLlegar, false, true) +
               formatearBloqueSeccion("📝 2. POR REGISTRAR", porRegistrar, true, false) +
               formatearBloqueSeccion("🔍 3. POR RECONOCER", porReconocer, false, false) +
               formatearBloqueSeccion("📋 4. POR VALIDAR", porValidar, false, false) +
               formatearBloqueSeccion("🚚 5. POR DESPACHAR", porDespachar, false, false);
    }

    // =========================================================================
    // 4. DETECCIÓN DE CAMPO Y VALOR A MODIFICAR (FECHA AUTOMÁTICA SI SE OMITE)
    // =========================================================================
    let campoAModificar = "";
    let nombreCampoLegible = "";
    let esIndicadorSi = false;
    let esCanal = false;
    let valorCanal = "";
    let esBorrar = textoLower.includes('borrar') || textoLower.includes('eliminar') || textoLower.includes('quitar');

    if (textoLower.includes('retirado bl') || textoLower.includes('retiro bl')) {
        campoAModificar = "retiro_bl";
        nombreCampoLegible = "Retiro de B/L";
        esIndicadorSi = true;
    } else if (textoLower.includes('prevaloración enviada') || textoLower.includes('prevaloracion enviada')) {
        campoAModificar = "prevaloracion_enviada";
        nombreCampoLegible = "Prevaloración Enviada";
        esIndicadorSi = true;
    } else if (textoLower.includes('valorado')) {
        campoAModificar = "documento_valorado";
        nombreCampoLegible = "Documento Valorado";
        esIndicadorSi = true;
    } else if (textoLower.includes('factura recibida') || textoLower.includes('recibida factura')) {
        campoAModificar = "recibida_factura";
        nombreCampoLegible = "Recibida Factura";
        esIndicadorSi = true;
    } else if (textoLower.includes('canal')) {
        campoAModificar = "canal";
        nombreCampoLegible = "Canal Aduanero";
        esCanal = true;
        if (textoLower.includes('verde')) valorCanal = "Verde";
        else if (textoLower.includes('amarillo')) valorCanal = "Amarillo";
        else if (textoLower.includes('rojo')) valorCanal = "Rojo";
    } else if (textoLower.includes('eta')) {
        campoAModificar = "eta_la_guaira";
        nombreCampoLegible = "ETA";
    } else if (textoLower.includes('llegada')) {
        campoAModificar = "fecha_llegada";
        nombreCampoLegible = "Fecha de Llegada";
    } else if (textoLower.includes('dai')) {
        campoAModificar = "fecha_registro_dai";
        nombreCampoLegible = "Registro DAI";
    } else if (textoLower.includes('transmitido') || textoLower.includes('transmisión') || textoLower.includes('transmision')) {
        campoAModificar = "fecha_transmision";
        nombreCampoLegible = "Fecha de Transmisión";
    } else if (textoLower.includes('reconocido') || textoLower.includes('reconocimiento')) {
        campoAModificar = "reconocimiento";
        nombreCampoLegible = "Reconocimiento";
    } else if (textoLower.includes('validado') || textoLower.includes('validación') || textoLower.includes('validacion')) {
        campoAModificar = "validacion";
        nombreCampoLegible = "Validación";
    } else if (textoLower.includes('despachado') || textoLower.includes('despacho')) {
        campoAModificar = "despacho";
        nombreCampoLegible = "Despacho";
    } else if (textoLower.includes('acta')) {
        campoAModificar = "retiro_acta";
        nombreCampoLegible = "Retiro de Acta";
    }

    if (campoAModificar && (textoLower.includes('actualizar') || textoLower.includes('actualiza') || esIndicadorSi || esCanal || esBorrar || textoLower.includes('transmitido') || textoLower.includes('reconocido') || textoLower.includes('validado') || textoLower.includes('despachado') || textoLower.includes('retirada') || textoLower.includes('canal') || textoLower.includes('valorado') || textoLower.includes('factura') || textoLower.includes('prevaloración'))) {
        
        let valorNuevo = "";
        if (esBorrar) {
            valorNuevo = "";
        } else if (esIndicadorSi) {
            valorNuevo = "SI";
        } else if (esCanal) {
            valorNuevo = valorCanal;
            if (!valorNuevo) return `❌ Debes indicar el color del canal (Verde, Amarillo o Rojo).`;
        } else {
            // Intenta buscar fecha en el texto; si no la encuentra, usa la fecha actual del día (DD/MM/YYYY)
            const matchFecha = textoMensaje.match(/\b(\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})\b/);
            if (matchFecha) {
                valorNuevo = matchFecha[0].replace(/-/g, '/');
            } else {
                const ahora = new Date();
                const dia = String(ahora.getDate()).padStart(2, '0');
                const mes = String(ahora.getMonth() + 1).padStart(2, '0');
                const anio = ahora.getFullYear();
                valorNuevo = `${dia}/${mes}/${anio}`;
            }
        }

        const recordsQuery = await fetch(`https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/expedientes`);
        const resultQuery = await recordsQuery.json();
        if (!resultQuery.documents) return `⚠️ No hay registros en Firestore.`;

        let expedienteEncontrado = null;
        for (let doc of resultQuery.documents) {
            let fields = doc.fields || {};
            let numExp = (fields.num_expediente?.stringValue || "").toLowerCase().replace(/[^a-z0-9]/g, '');
            let blNum = (fields.awb_bl?.stringValue || fields.bl?.stringValue || fields.numero_bl?.stringValue || "").toLowerCase().replace(/[^a-z0-9]/g, '');
            
            let limpioTexto = textoLower.replace(/[^a-z0-9]/g, '');

            if ((numExp.length > 1 && limpioTexto.includes(numExp)) || (blNum.length > 2 && limpioTexto.includes(blNum))) {
                expedienteEncontrado = { 
                    id: doc.name.split('/').pop(), 
                    fields, 
                    numExpediente: fields.num_expediente?.stringValue || 'S/N',
                    cliente: fields.cliente?.stringValue || 'S/C',
                    bl: fields.awb_bl?.stringValue || fields.bl?.stringValue || fields.numero_bl?.stringValue || 'S/N'
                };
                break;
            }
        }

        if (!expedienteEncontrado) return `⚠️ No se encontró el expediente o B/L indicado en el mensaje.`;

        const updateUrl = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/expedientes/${expedienteEncontrado.id}?updateMask.fieldPaths=${campoAModificar}`;
        const patchBody = {
            fields: {
                ...expedienteEncontrado.fields,
                [campoAModificar]: { stringValue: valorNuevo }
            }
        };

        const patchRes = await fetch(updateUrl, {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(patchBody)
        });

        if (patchRes.ok) {
            if (esBorrar) {
                return `🗑️ *¡CAMPO ELIMINADO CON ÉXITO!* ⚠️\n\n📂 Exp: *${expedienteEncontrado.numExpediente}*\n👤 Cliente: *${expedienteEncontrado.cliente}*\n📦 B/L: *${expedienteEncontrado.bl}*\n⚙️ Campo: *${nombreCampoLegible}* (Ha quedado vacío).`;
            } else if (esIndicadorSi) {
                return `✅ *¡INDICADOR ACTUALIZADO A SÍ!* 🚀\n\n📂 Exp: *${expedienteEncontrado.numExpediente}*\n👤 Cliente: *${expedienteEncontrado.cliente}*\n📦 B/L: *${expedienteEncontrado.bl}*\n⚙️ Campo: *${nombreCampoLegible}* ➔ *SÍ*`;
            } else if (esCanal) {
                return `🎨 *¡CANAL ACTUALIZADO CON ÉXITO!* 🚀\n\n📂 Exp: *${expedienteEncontrado.numExpediente}*\n👤 Cliente: *${expedienteEncontrado.cliente}*\n📦 B/L: *${expedienteEncontrado.bl}*\n⚙️ Campo: *Canal Aduanero* ➔ *${valorNuevo.toUpperCase()}*`;
            } else {
                return `✅ *¡FECHA ACTUALIZADA CON ÉXITO!* 🚀\n\n📂 Exp: *${expedienteEncontrado.numExpediente}*\n👤 Cliente: *${expedienteEncontrado.cliente}*\n📦 B/L: *${expedienteEncontrado.bl}*\n⚙️ Campo: *${nombreCampoLegible}*\n📅 Fecha Actualizada: *${formatearFechaDDMMYYYY(valorNuevo)}*`;
            }
        } else {
            return `❌ Error al actualizar en Firestore.`;
        }
    }

    return null;
}

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
app.get('/api/whatsapp/estado', (req, res) => {
    res.json({ success: true, estado: estadoConexionWhatsapp, ready: whatsappClientReady, qrCodeDataUrl: ultimoQrGenerado });
});

// =========================================================================
// RUTA DE FORZAR ALERTAS (MULTICANAL)
// =========================================================================
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
    return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
}

async function enviarAlertaMulticanal(listaAlertas, tituloReporte, canalesForzados = null) {
    try {
        let enviarTg = canalesForzados ? canalesForzados.telegram : Boolean(TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID);
        let enviarWa = canalesForzados ? canalesForzados.whatsapp : Boolean(whatsappClientReady && sock);

        if (enviarTg && TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID) {
            let mensajesEnumerados = listaAlertas.map((item, index) => `${index + 1}️⃣ ${item}`);
            let mensajeConsolidado = `🚨 *LOGISTATUS PRO - ${tituloReporte}* 🚨\n\n` + mensajesEnumerados.join("\n\n. . . . . . . . . . . . . . . . . . . . . . . .\n\n");
            await enviarAlertaTelegramServidorPersonalizado(TELEGRAM_CHAT_ID, mensajeConsolidado);
        }

        if (enviarWa && whatsappClientReady && sock) {
            const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us";
            let mensajesEnumerados = listaAlertas.map((item, index) => `${index + 1}️⃣ ${item}`);
            let mensajeConsolidado = `🚨 *LOGISTATUS PRO - ${tituloReporte}* 🚨\n\n` + mensajesEnumerados.join("\n\n. . . . . . . . . . . . . . . . . . . . . . . .\n\n");
            await sock.sendMessage(GRUPO_AUTORIZADO_JID, { text: mensajeConsolidado });
        }
    } catch (err) {
        console.error("❌ Error en envío multicanal:", err);
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

// ==========================================
// 3. TAREAS PROGRAMADAS (CRON)
// ==========================================
// A. Correos automáticos: Lunes a Viernes a las 10:00 AM y 6:00 PM
cron.schedule('0 10,18 * * 1-5', async () => {
    try {
        console.log("⏰ [CRON CORREO] Ejecutando envío automático de correo...");
        const correoCliente = process.env.CORREO_DESTINATARIO_AUTOMATICO;
        if (!correoCliente) return;
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) return;
    } catch (error) {
        console.error("❌ [CRON CORREO] Error:", error);
    }
}, { timezone: "America/Caracas" });

// B. Alertas Multicanal (WhatsApp y Telegram): Todos los días a las 9, 12, 15, 18, 19, 20 horas
cron.schedule('0 9,12,15,18,19,20 * * *', async () => {
    try {
        console.log("⏰ [CRON MULTICANAL] Ejecutando reporte automático de alertas...");
        const alertasGeneradas = await generarListaAlertasSistema();
        if (alertasGeneradas && alertasGeneradas.length > 0) {
            await enviarAlertaMulticanal(alertasGeneradas, "REPORTE AUTOMÁTICO PROGRAMADO");
        }
    } catch (error) {
        console.error("❌ [CRON MULTICANAL] Error en reporte automático:", error);
    }
}, { timezone: "America/Caracas" });

// ==========================================
// 4. INICIO DEL SERVIDOR
// ==========================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => { console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`); });