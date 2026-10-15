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
const FIREBASE_PROJECT_ID = process.env.FIREBASE_PROJECT_ID || "statusylogistica";[cite: 1]
const FIREBASE_API_KEY = process.env.FIREBASE_API_KEY || "";[cite: 1]

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
        const { state, saveCreds } = await useMultiFileAuthState(path.join(__dirname, 'auth_info_baileys'));[cite: 1]
        
        sock = makeWASocket({
            auth: state,
            printQRInTerminal: false
        });

        sock.ev.on('connection.update', async (update) => {
            const { connection, lastDisconnect, qr } = update;

            if (qr) {
                console.log('📱 [WHATSAPP] Nuevo código QR generado.');[cite: 1]
                estadoConexionWhatsapp = "ESPERANDO_ESCANEO";
                try {
                    ultimoQrGenerado = await qrcode.toDataURL(qr);
                } catch (err) {
                    console.error("❌ Error al convertir QR:", err);[cite: 1]
                }
            }

            if (connection === 'open') {
                console.log('✅ [WHATSAPP] ¡Conectado y listo!');[cite: 1]
                whatsappClientReady = true;
                estadoConexionWhatsapp = "CONECTADO";
                ultimoQrGenerado = "";

                try {
                    const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us";[cite: 1]
                    await sock.sendMessage(GRUPO_AUTORIZADO_JID, { text: "🟢 *El bot de aduanas se ha conectado y está operativo.*" });[cite: 1]
                } catch (e) { console.error("No se pudo enviar aviso de conexión", e); }[cite: 1]
            }

            if (connection === 'close') {
                const statusCode = lastDisconnect?.error?.output?.statusCode;
                const shouldReconnect = statusCode !== DisconnectReason.loggedOut;
                
                console.log(`⚠️ [WHATSAPP] Conexión cerrada. Código: ${statusCode}. Reconectando: ${shouldReconnect}`);[cite: 1]
                whatsappClientReady = false;
                estadoConexionWhatsapp = "DESCONECTADO";

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
                if (m.type !== 'notify') return;[cite: 1]

                const msg = m.messages[0];
                if (!msg.message || msg.key.fromMe) return;[cite: 1]

                const remitente = msg.key.remoteJid;
                const textoMensaje = (msg.message.conversation || msg.message.extendedTextMessage?.text || "").trim();
                const textoLower = textoMensaje.toLowerCase();

                console.log(`📩 [MENSAJE RECIBIDO] De: ${remitente} | Texto: ${textoMensaje}`);[cite: 1]

                const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us";[cite: 1] 
                
                if (remitente !== GRUPO_AUTORIZADO_JID) {
                    console.log(`⛔ Mensaje ignorado por seguridad (No proviene del grupo autorizado).`);[cite: 1]
                    return;
                }

                // =========================================================================
                // 📈 COMANDO RÁPIDO: RESUMEN Y CONTEO DE EXPEDIENTES POR ESTADO
                // =========================================================================
                if (textoLower.includes('resumen') || textoLower.includes('conteo') || textoLower.includes('totales')) {
                    console.log(`📊 [WHATSAPP] Solicitud de resumen por estados detectada.`);
                    
                    const records = await obtenerExpedientesFirestore();
                    if (!records || records.length === 0) {
                        await sock.sendMessage(remitente, { text: "✅ No hay registros en Firestore para resumir." });
                        return;
                    }

                    let porEta = [];
                    let porRegistrar = [];
                    let porReconocer = [];
                    let porValidar = [];
                    let porDespachar = [];

                    for (const rec of records) {
                        if (rec.despacho) continue;

                        const fechaEtaRaw = rec.fecha_llegada || rec.eta_la_guaira || '';
                        const hasLlegada = Boolean(rec.fecha_llegada && rec.fecha_llegada.trim() !== "");
                        const hasDai = Boolean(rec.fecha_registro_dai && rec.fecha_registro_dai.trim() !== "");
                        const hasTransmision = Boolean(rec.fecha_transmision && rec.fecha_transmision.trim() !== "");
                        const hasReconocimiento = Boolean(rec.reconocimiento && rec.reconocimiento.trim() !== "");
                        const hasValidacion = Boolean(rec.validacion && rec.validacion.trim() !== "");

                        const dArribo = parseFechaLocalServidor(fechaEtaRaw);
                        const today = new Date();
                        today.setHours(0,0,0,0);
                        if (dArribo && !hasLlegada) {
                            porEta.push({ rec, fechaObj: dArribo, fechaStr: formatearFechaDDMMYYYY(fechaEtaRaw) });
                        }

                        if (!hasDai) {
                            porRegistrar.push(rec);
                        }

                        if (hasTransmision && !hasReconocimiento) {
                            porReconocer.push(rec);
                        }

                        if (hasReconocimiento && !hasValidacion) {
                            porValidar.push(rec);
                        }

                        if (hasValidacion) {
                            porDespachar.push(rec);
                        }
                    }

                    porEta.sort((a, b) => a.fechaObj - b.fechaObj);
                    let infoEtaMasProximo = porEta.length > 0 ? ` (El más próximo: ${porEta[0].fechaStr})` : "";

                    let mensajeResumen = `📊 *RESUMEN DE EXPEDIENTES - LOGISTATUS PRO* 📊\n\n` +
                        `🚢 *Por ETA:* ${porEta.length}${infoEtaMasProximo}\n` +
                        `📝 *Por Registrar (DAI):* ${porRegistrar.length}\n` +
                        `🔍 *Por Reconocer:* ${porReconocer.length}\n` +
                        `📋 *Por Validar:* ${porValidar.length}\n` +
                        `🚀 *Por Despachar:* ${porDespachar.length}`;

                    await sock.sendMessage(remitente, { text: mensajeResumen });
                    return;
                }
                // =========================================================================

                let campoAModificar = "";
                let nombreCampoLegible = "";
                
                if (textoLower.includes('eta')) {
                    campoAModificar = "eta_la_guaira";
                    nombreCampoLegible = "ETA La Guaira";
                } else if (textoLower.includes('llegada')) {
                    campoAModificar = "fecha_llegada";
                    nombreCampoLegible = "Fecha de Llegada";
                } else if (textoLower.includes('dai')) {
                    campoAModificar = "fecha_registro_dai";
                    nombreCampoLegible = "Registro DAI";
                } else if (textoLower.includes('transmision') || textoLower.includes('transmitido')) {
                    campoAModificar = "fecha_transmision";
                    nombreCampoLegible = "Fecha de Transmisión";
                } else if (textoLower.includes('reconocimiento')) {
                    campoAModificar = "reconocimiento";
                    nombreCampoLegible = "Reconocimiento";
                } else if (textoLower.includes('validacion')) {
                    campoAModificar = "validacion";
                    nombreCampoLegible = "Validación";
                } else if (textoLower.includes('despacho')) {
                    campoAModificar = "despacho";
                    nombreCampoLegible = "Despacho";
                } else if (textoLower.includes('acta')) {
                    campoAModificar = "recibida_acta_recepcion";
                    nombreCampoLegible = "Acta de Recepción";
                }

                // =========================================================================
                // 📊 COMANDO RÁPIDO: SOLICITAR ALERTAS ACTIVAS DESDE EL GRUPO
                // =========================================================================
                if (textoLower.includes('alertas') || textoLower.includes('reporte') || textoLower.includes('estatus')) {
                    console.log(`📊 [WHATSAPP] Solicitud de alertas manual detectada en el grupo.`);[cite: 1]
                    
                    const records = await obtenerExpedientesFirestore();
                    if (!records || records.length === 0) {
                        await sock.sendMessage(remitente, { text: "✅ No hay registros en Firestore para evaluar." });[cite: 1]
                        return;
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

                        if (dArriboObj && !hasLlegadaRegistrada) {
                            const diffLlegadaDays = Math.round((dArriboObj - today) / (1000 * 60 * 60 * 24));
                            if (diffLlegadaDays >= 0 && diffLlegadaDays <= 5) {
                                const textoDias = diffLlegadaDays === 0 ? "¡LLEGA HOY!" : `Faltan ${diffLlegadaDays}d`;
                                alertasEta.push(`${iconoVia} *${destinoTxt.toUpperCase()}*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ *${textoDias}* (${fechaEtaFormatted})`);
                            }
                        }

                        if (dArriboObj && !hasLlegadaRegistrada && !hasRegistroDai) {
                            const diffEtaDaiDays = Math.round((dArriboObj - today) / (1000 * 60 * 60 * 24));
                            if (diffEtaDaiDays >= 0 && diffEtaDaiDays <= 5) {
                                const textoEtaDai = diffEtaDaiDays === 0 ? "¡LLEGA HOY!" : `Faltan ${diffEtaDaiDays}d para el ETA`;
                                alertasDaiPendientes.push(`⚠️ *DAI POR REGISTRAR (ETA PRÓXIMO)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoEtaDai} (${fechaEtaFormatted})\n⚠️ *La fecha de registro DAI se encuentra en blanco.*`);
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
                            const fechaPlazoStr = formatearFechaDDMMYYYY(dDeadline.toISOString().split('T')[0]);

                            if (diffPlazoDays >= 0 && diffPlazoDays <= 2) {
                                const textoPlazo = diffPlazoDays === 0 ? "¡VENCE HOY EL PLAZO!" : `Quedan ${diffPlazoDays} día(s) de plazo`;
                                alertasVencimientos.push(`⚠️ *PLAZO DE DECLARACIÓN PRÓXIMO A VENCER (Art. 47 LOA)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoPlazo} (Límite: ${fechaPlazoStr}) - *Riesgo Multa Art. 177 Num. 7 de la L.O.A.*`);
                            } else if (diffPlazoDays < 0) {
                                const diasExt = Math.abs(diffPlazoDays);
                                alertasVencimientos.push(`🚨 *EXTEMPORANEIDAD - ART. 47*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ El plazo de 5 días hábiles venció el ${fechaPlazoStr} (Han transcurrido ${diasExt} día(s) en extemporaneidad).\n🚨 *¡Aplica Multa Art. 177 Num 7 de la L.O.A.! (Persiste hasta transmitir)*`);
                            }
                        }

                        if (rec.culminacion_dias_libres) {
                            const culmDate = parseFechaLocalServidor(rec.culminacion_dias_libres);
                            const culmFormatted = formatearFechaDDMMYYYY(rec.culminacion_dias_libres);
                            if (culmDate) {
                                const diffDaysAlm = Math.round((culmDate - today) / (1000 * 60 * 60 * 24));
                                if (diffDaysAlm >= 0 && diffDaysAlm <= 2) {
                                    const textoAlm = diffDaysAlm === 0 ? "VENCE HOY" : `Faltan ${diffDaysAlm}d`;
                                    alertasVencimientos.push(`⚠️ *ALMACÉN PRÓXIMO A VENCER*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoAlm} (${culmFormatted})`);
                                } else if (diffDaysAlm < 0) {
                                    alertasVencimientos.push(`🚨 *ALMACÉN VENCIDO*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⚠️ Venció hace ${Math.abs(diffDaysAlm)} día(s) el ${culmFormatted}`);
                                }
                            }
                        }

                        const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
                        const vencDaiFormatted = formatearFechaDDMMYYYY(rec.vencimiento_dai);
                        if (!hasLlegadaRegistrada && !hasRegistroDai && dVencimientoDai) {
                            const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                            if (diffVencDays < 0) {
                                const diasVencida = Math.abs(diffVencDays);
                                alertasVencimientos.push(`🚨 *DAI VENCIDA (SIN REGISTRAR)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⚠️ Venció el ${vencDaiFormatted} (${diasVencida} día(s) atrás) sin registrar DAI.\n🚨 *¡Aplica Multa Art. 177 Num. 9 de la L.O.A.! (Persiste hasta registrar)*`);
                            } else if (diffVencDays >= 0 && diffVencDays <= 4) {
                                const textoCountdown = diffVencDays === 0 ? "¡VENCE HOY!" : `Faltan ${diffVencDays}d`;
                                alertasVencimientos.push(`⚠️ *DAI PRÓXIMA A VENCER*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoCountdown} (${vencDaiFormatted})\n🚢 ETA: *${fechaEtaFormatted}* - *Riesgo Multa Art 177 Num 9 de la L.O.A*`);
                            }
                        }
                    }

                    let listaFinalOrdenada = [...alertasEta, ...alertasDaiPendientes, ...alertasVencimientos];
                    if (listaFinalOrdenada.length === 0) {
                        await sock.sendMessage(remitente, { text: "✅ ¡Excelente! No hay alertas críticas pendientes en este momento." });[cite: 1]
                        return;
                    }

                    let mensajesEnumerados = listaFinalOrdenada.map((item, index) => `${index + 1}️⃣ ${item}`);
                    let mensajeConsolidado = "🚨 *LOGISTATUS PRO - REPORTE SOLICITADO* 🚨\n\n" + mensajesEnumerados.join("\n\n. . . . . . . . . . . . . . . . . . . . . . . .\n\n");[cite: 1]
                    
                    await sock.sendMessage(remitente, { text: mensajeConsolidado });[cite: 1]
                    return;
                }
                // =========================================================================

                if (!campoAModificar) return;

                const pideBorrar = textoLower.includes('borrar') || textoLower.includes('eliminar') || textoLower.includes('vaciar');
                let fechaDetectada = "";
                
                if (!pideBorrar) {
                    const matchFecha = textoMensaje.match(/\b(\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})\b/);
                    fechaDetectada = matchFecha ? matchFecha[0].replace(/-/g, '/') : new Date().toLocaleDateString("es-VE", { timeZone: "America/Caracas" });[cite: 1]
                }

                const projectId = process.env.FIREBASE_PROJECT_ID || "statusylogistica";[cite: 1]
                const urlQuery = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/expedientes`;[cite: 1]
                const response = await fetch(urlQuery);
                const result = await response.json();

                if (!result.documents || result.documents.length === 0) return;[cite: 1]

                let listaExpedientes = result.documents.map(doc => {
                    let fields = doc.fields || {};
                    let recParsed = {};
                    for (let key in fields) {
                        recParsed[key] = fields[key].stringValue || fields[key].doubleValue || fields[key].integerValue || fields[key].booleanValue || "";[cite: 1]
                    }
                    return {
                        id: doc.name.split('/').pop(),
                        fields: fields,
                        raw: recParsed,
                        num_expediente: (recParsed.num_expediente || "").toLowerCase().trim(),
                        bl: (recParsed.awb_bl || recParsed.bl || recParsed.numero_bl || "").toLowerCase().trim()
                    };
                });

                let expedienteEncontrado = null;
                for (let item of listaExpedientes) {
                    let limpioTexto = textoLower.replace(/[^a-z0-9]/g, '');
                    let limpioExp = item.num_expediente.replace(/[^a-z0-9]/g, '');
                    let limpioBl = item.bl.replace(/[^a-z0-9]/g, '');

                    if ((limpioExp.length > 1 && limpioTexto.includes(limpioExp)) ||
                        (limpioBl.length > 2 && limpioTexto.includes(limpioBl))) {
                        expedienteEncontrado = item;
                        break;
                    }
                }

                if (!expedienteEncontrado) {
                    await sock.sendMessage(remitente, { text: `⚠️ No se encontró el expediente indicado en el mensaje.` });[cite: 1]
                    return;
                }

                let datosActuales = expedienteEncontrado.raw;
                let numExpReal = datosActuales.num_expediente || 'S/N';
                let clientReal = datosActuales.cliente || 'S/C';
                let blReal = datosActuales.awb_bl || datosActuales.bl || 'S/N';

                if (!pideBorrar) {
                    if (campoAModificar === 'fecha_transmision' && (!datosActuales.fecha_llegada && !datosActuales.eta_la_guaira)) {
                        await sock.sendMessage(remitente, { text: `❌ *ACCIÓN DENEGADA*\n\n📂 Exp: *${numExpReal}*\n👤 Cliente: *${clientReal}*\n\nNo se puede registrar la *Transmisión* porque la **Fecha de Llegada / ETA** previa se encuentra en blanco.` });[cite: 1]
                        return;
                    }
                    if (campoAModificar === 'reconocimiento' && !datosActuales.fecha_transmision) {
                        await sock.sendMessage(remitente, { text: `❌ *ACCIÓN DENEGADA*\n\n📂 Exp: *${numExpReal}*\n👤 Cliente: *${clientReal}*\n\nNo se puede registrar el *Reconocimiento* porque la **Fecha de Transmisión** inmediatamente anterior no está cargada.` });[cite: 1]
                        return;
                    }
                    if (campoAModificar === 'validacion' && !datosActuales.reconocimiento) {
                        await sock.sendMessage(remitente, { text: `❌ *ACCIÓN DENEGADA*\n\n📂 Exp: *${numExpReal}*\n👤 Cliente: *${clientReal}*\n\nNo se puede registrar la *Validación* porque el **Reconocimiento** previo no está registrado.` });[cite: 1]
                        return;
                    }
                    if (campoAModificar === 'despacho' && !datosActuales.validacion) {
                        await sock.sendMessage(remitente, { text: `❌ *ACCIÓN DENEGADA*\n\n📂 Exp: *${numExpReal}*\n👤 Cliente: *${clientReal}*\n\nNo se puede registrar el *Despacho* porque la **Validación** inmediatamente anterior no está cargada.` });[cite: 1]
                        return;
                    }
                }

                const updateUrl = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/expedientes/${expedienteEncontrado.id}?updateMask.fieldPaths=${campoAModificar}`;[cite: 1]
                
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
                    let mensajeRespuesta = "";
                    let fechaDdmmyyyyLog = formatearFechaDDMMYYYY(fechaDetectada);
                    if (pideBorrar || !fechaDetectada) {
                        mensajeRespuesta = `🗑️ *¡FECHA ELIMINADA EXITOSAMENTE!* ⚠️\n\n📂 Exp: *${numExpReal}*\n👤 Cliente: *${clientReal}*\n📦 B/L: *${blReal}*\n⚙️ Campo: *${nombreCampoLegible}*\n📌 _Fecha eliminada según su petición._`;[cite: 1]
                    } else {
                        mensajeRespuesta = `✅ *¡FECHA ACTUALIZADA CON ÉXITO!* 🚀\n\n📂 Exp: *${numExpReal}*\n👤 Cliente: *${clientReal}*\n📦 B/L: *${blReal}*\n⚙️ Campo: *${nombreCampoLegible}*\n📅 Fecha Registrada: *${fechaDdmmyyyyLog}*\n📌 _Fecha actualizada según su petición._`;[cite: 1]
                    }
                    await sock.sendMessage(remitente, { text: mensajeRespuesta });[cite: 1]
                } else {
                    const errData = await patchRes.json();
                    await sock.sendMessage(remitente, { text: `❌ Error al actualizar en Firestore: ${errData.error?.message || 'Desconocido'}` });[cite: 1]
                }

            } catch (err) {
                console.error("❌ Error procesando mensaje en Baileys:", err);[cite: 1]
            }
        });

    } catch (err) {
        console.error("❌ Error al iniciar Baileys en la nube:", err);[cite: 1]
    }
}

iniciarBaileysWhatsApp();[cite: 1]

// ==========================================
// 1. RUTAS ESENCIALES Y DE CONFIGURACIÓN
// ==========================================
app.get('/api/config', (req, res) => {
    res.json({
        firebaseApiKey: FIREBASE_API_KEY,[cite: 1]
        firebaseProjectId: FIREBASE_PROJECT_ID,[cite: 1]
        telegramBotToken: TELEGRAM_BOT_TOKEN || '',[cite: 1]
        telegramChatId: TELEGRAM_CHAT_ID || ''[cite: 1]
    });
});

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));[cite: 1]
});

app.get('/health', (req, res) => {
    res.status(200).send('LOGISTATUS PRO - Servidor Activo OK');[cite: 1]
});

app.get('/api/whatsapp/estado', (req, res) => {
    res.json({
        success: true,
        estado: estadoConexionWhatsapp,[cite: 1]
        ready: whatsappClientReady,[cite: 1]
        qrCodeDataUrl: ultimoQrGenerado[cite: 1]
    });
});

app.post('/api/whatsapp/enviar', async (req, res) => {
    const { telefono, mensaje } = req.body;
    if (!whatsappClientReady || !sock) {
        return res.status(400).json({ success: false, message: "El bot de WhatsApp aún no está conectado." });[cite: 1]
    }
    if (!telefono || !mensaje) {
        return res.status(400).json({ success: false, message: "Faltan el número de teléfono o el mensaje." });[cite: 1]
    }

    try {
        let numeroLimpios = telefono.replace(/\D/g, '');
        const jid = `${numeroLimpios}@s.whatsapp.net`;

        await sock.sendMessage(jid, { text: mensaje });[cite: 1]
        res.json({ success: true, message: "Mensaje de WhatsApp enviado con éxito." });[cite: 1]
    } catch (error) {
        console.error("❌ Error al enviar WhatsApp con Baileys:", error);[cite: 1]
        res.status(500).json({ success: false, message: error.message });[cite: 1]
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

async function enviarReporteWhatsAppGrupo(listaAlertas, tituloReporte) {
    try {
        if (!whatsappClientReady || !sock) return;
        const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us";[cite: 1]
        let mensajesEnumerados = listaAlertas.map((item, index) => `${index + 1}️⃣ ${item}`);
        let mensajeConsolidado = `🚨 *LOGISTATUS PRO - ${tituloReporte}* 🚨\n\n` + mensajesEnumerados.join("\n\n. . . . . . . . . . . . . . . . . . . . . . . .\n\n");
        await sock.sendMessage(GRUPO_AUTORIZADO_JID, { text: mensajeConsolidado });[cite: 1]
    } catch (err) {
        console.error("❌ Error al enviar reporte al grupo de WhatsApp:", err);[cite: 1]
    }
}

async function enviarCorreoSistema(opcionesMail) {
    try {
        const apiKey = process.env.RESEND_API_KEY; 
        const remitente = process.env.RESEND_REMITE || "onboarding@resend.dev";[cite: 1]

        if (!apiKey || apiKey.includes("TU_API_KEY")) {
            throw new Error("Falta configurar la RESEND_API_KEY.");[cite: 1]
        }

        let destinatarios = opcionesMail.to;
        if (!Array.isArray(destinatarios)) destinatarios = [destinatarios];
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
            from: `LOGISTATUS PRO <${remitente}>`,[cite: 1]
            to: destinatarios,
            subject: opcionesMail.subject || 'Notificación LOGISTATUS PRO',[cite: 1]
            html: opcionesMail.html || opcionesMail.text || 'Sin contenido'[cite: 1]
        };

        if (attachmentsFormatted.length > 0) payload.attachments = attachmentsFormatted;

        const response = await fetch('https://api.resend.com/emails', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
            body: JSON.stringify(payload)
        });

        const data = await response.json();
        if (!response.ok) throw new Error(`Error API: ${JSON.stringify(data)}`);[cite: 1]
        return { success: true };
    } catch (error) {
        console.error("❌ Error al enviar el correo:", error);[cite: 1]
        throw error;
    }
}

async function enviarAlertaTelegramServidor(mensajeTexto) {
    try {
        const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;[cite: 1]
        await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ chat_id: TELEGRAM_CHAT_ID, text: mensajeTexto, parse_mode: 'Markdown' })[cite: 1]
        });
    } catch (err) {
        console.error("❌ Error Telegram:", err);[cite: 1]
    }
}

async function obtenerExpedientesFirestore() {
    try {
        const projectId = process.env.FIREBASE_PROJECT_ID;[cite: 1]
        const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/expedientes`;[cite: 1]
        const response = await fetch(url);
        const result = await response.json();
        
        if (!result.documents) return [];[cite: 1]

        return result.documents.map(doc => {
            let fields = doc.fields || {};
            let parsedRecord = {};
            for (let key in fields) {
                let valObj = fields[key];
                parsedRecord[key] = valObj.stringValue || valObj.doubleValue || valObj.integerValue || valObj.booleanValue || "";[cite: 1]
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
        if (p.length === 3 && p[2].length === 4) return new Date(`${p[2]}-${p[1]}-${p[0]}T00:00:00`);[cite: 1]
    } else if (str.includes('-')) {
        const p = str.split('-');
        if (p.length === 3 && p[0].length === 4) return new Date(`${p[0]}-${p[1]}-${p[2]}T00:00:00`);[cite: 1]
    }
    return null;
}

// ==========================================
// 3. TAREAS PROGRAMADAS (CRON)
// ==========================================
cron.schedule('0 10,18 * * *', async () => {
    console.log("⏰ [CRON] Ejecutando generación y envío automático...");[cite: 1]
    try {
        const correoCliente = process.env.CORREO_DESTINATARIO_AUTOMATICO;[cite: 1]
        if (!correoCliente) return;
        const destinatariosAutomaticos = [correoCliente];[cite: 1]
        
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
            top: { style: "thin", color: { rgb: "FFFFFF" } },[cite: 1]
            bottom: { style: "thin", color: { rgb: "FFFFFF" } },[cite: 1]
            left: { style: "thin", color: { rgb: "FFFFFF" } },[cite: 1]
            right: { style: "thin", color: { rgb: "FFFFFF" } }[cite: 1]
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
                cellA.s = { font: { name: "Segoe UI", sz: 11, bold: true, color: { rgb: "1E3A8A" } } };[cite: 1]
                continue;
            }

            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                if (!wsMaster[cellAddress]) continue;

                let cellStyle = { border: blackBorder, alignment: { vertical: "center", horizontal: "center", wrapText: true }, font: { name: "Segoe UI", sz: 10 } };[cite: 1]
                const headerCheck = wsMaster[XLSX.utils.encode_cell({ r: R, c: 0 })];
                if (headerCheck && excelHeaders.includes(headerCheck.v)) {
                    cellStyle.fill = { fgColor: { rgb: "1E3A8A" } };[cite: 1]
                    cellStyle.font = { name: "Segoe UI", sz: 10, bold: true, color: { rgb: "FFFFFF" } };[cite: 1]
                } else {
                    cellStyle.font = { name: "Segoe UI", sz: 10, color: { rgb: "000000" }, bold: true };[cite: 1]
                    cellStyle.fill = { fgColor: { rgb: currentSection === 3 ? "FEE2E2" : currentSection === 2 ? "D9E1F2" : "F8FAFC" } };[cite: 1]
                }
                wsMaster[cellAddress].s = cellStyle;
            }
        }

        XLSX.utils.book_append_sheet(wb, wsMaster, "Reporte Consolidado");
        const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
        const excelBase64String = Buffer.from(excelBuffer).toString('base64');

        await enviarCorreoSistema({
            to: [correoCliente],[cite: 1]
            subject: "📊 Reporte Maestro Automático - LOGISTATUS PRO",[cite: 1]
            text: "Adjunto el reporte automático consolidado.",[cite: 1]
            attachments: [{ filename: 'Reporte_Automatico.xlsx', content: excelBase64String, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }]
        });
    } catch (error) {
        console.error("❌ [CRON] Error al enviar reporte automático:", error);[cite: 1]
    }
}, { timezone: "America/Caracas" });

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

            if (dArriboObj && !hasLlegadaRegistrada) {
                const diffLlegadaDays = Math.round((dArriboObj - today) / (1000 * 60 * 60 * 24));
                if (diffLlegadaDays >= 0 && diffLlegadaDays <= 5) {
                    const textoDias = diffLlegadaDays === 0 ? "¡LLEGA HOY!" : `Faltan ${diffLlegadaDays}d`;
                    alertasEta.push(`${iconoVia} *${destinoTxt.toUpperCase()}*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ *${textoDias}* (${fechaEtaFormatted})`);
                }
            }

            if (dArriboObj && !hasLlegadaRegistrada && !hasRegistroDai) {
                const diffEtaDaiDays = Math.round((dArriboObj - today) / (1000 * 60 * 60 * 24));
                if (diffEtaDaiDays >= 0 && diffEtaDaiDays <= 5) {
                    const textoEtaDai = diffEtaDaiDays === 0 ? "¡LLEGA HOY!" : `Faltan ${diffEtaDaiDays}d para el ETA`;
                    alertasDaiPendientes.push(`⚠️ *DAI POR REGISTRAR (ETA PRÓXIMO)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoEtaDai} (${fechaEtaFormatted})\n⚠️ *La fecha de registro DAI se encuentra en blanco.*`);
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
                    if (diaSemana !== 0 && diaSemana !== 6) diasHabilesAgregados++;
                    intentosMaximos++;
                }
                const diffPlazoDays = Math.round((dDeadline - today) / (1000 * 60 * 60 * 24));
                const fechaPlazoStr = formatearFechaDDMMYYYY(dDeadline.toISOString().split('T')[0]);

                if (diffPlazoDays >= 0 && diffPlazoDays <= 2) {
                    const textoPlazo = diffPlazoDays === 0 ? "¡VENCE HOY EL PLAZO!" : `Quedan ${diffPlazoDays} día(s) de plazo`;
                    alertasVencimientos.push(`⚠️ *PLAZO DE DECLARACIÓN PRÓXIMO A VENCER (Art. 47 LOA)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoPlazo} (Límite: ${fechaPlazoStr}) - *Riesgo Multa Art. 177 Num. 7 de la L.O.A.*`);
                } else if (diffPlazoDays < 0) {
                    const diasExt = Math.abs(diffPlazoDays);
                    alertasVencimientos.push(`🚨 *EXTEMPORANEIDAD - ART. 47*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ El plazo de 5 días hábiles venció el ${fechaPlazoStr} (Han transcurrido ${diasExt} día(s) en extemporaneidad).\n🚨 *¡Aplica Multa Art. 177 Num 7 de la L.O.A.!*`);
                }
            }

            if (rec.culminacion_dias_libres) {
                const culmDate = parseFechaLocalServidor(rec.culminacion_dias_libres);
                const culmFormatted = formatearFechaDDMMYYYY(rec.culminacion_dias_libres);
                if (culmDate) {
                    const diffDaysAlm = Math.round((culmDate - today) / (1000 * 60 * 60 * 24));
                    if (diffDaysAlm >= 0 && diffDaysAlm <= 2) {
                        const textoAlm = diffDaysAlm === 0 ? "VENCE HOY" : `Faltan ${diffDaysAlm}d`;
                        alertasVencimientos.push(`⚠️ *ALMACÉN PRÓXIMO A VENCER*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoAlm} (${culmFormatted})`);
                    } else if (diffDaysAlm < 0) {
                        alertasVencimientos.push(`🚨 *ALMACÉN VENCIDO*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⚠️ Venció hace ${Math.abs(diffDaysAlm)} día(s) el ${culmFormatted}`);
                    }
                }
            }

            const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
            const vencDaiFormatted = formatearFechaDDMMYYYY(rec.vencimiento_dai);
            if (!hasLlegadaRegistrada && !hasRegistroDai && dVencimientoDai) {
                const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                if (diffVencDays < 0) {
                    const diasVencida = Math.abs(diffVencDays);
                    alertasVencimientos.push(`🚨 *DAI VENCIDA (SIN REGISTRAR)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⚠️ Venció el ${vencDaiFormatted} (${diasVencida} día(s) atrás) sin registrar DAI.\n🚨 *¡Aplica Multa Art. 177 Num. 9 de la L.O.A.!*`);
                } else if (diffVencDays >= 0 && diffVencDays <= 4) {
                    const textoCountdown = diffVencDays === 0 ? "¡VENCE HOY!" : `Faltan ${diffVencDays}d`;
                    alertasVencimientos.push(`⚠️ *DAI PRÓXIMA A VENCER*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoCountdown} (${vencDaiFormatted})\n🚢 ETA: *${fechaEtaFormatted}* - *Riesgo Multa Art 177 Num 9 de la L.O.A*`);
                }
            }
        }

        let listaFinalOrdenada = [...alertasEta, ...alertasDaiPendientes, ...alertasVencimientos];
        if (listaFinalOrdenada.length > 0) {
            let mensajeConsolidado = "🚨 *LOGISTATUS PRO - REPORTE AUTOMÁTICO CRON* 🚨\n\n" + listaFinalOrdenada.map((item, index) => `${index + 1}️⃣ ${item}`).join("\n\n. . . . . . . . . . . . . . . . . . . . . . . .\n\n");
            
            await enviarAlertaTelegramServidor(mensajeConsolidado);
            await enviarReporteWhatsAppGrupo(listaFinalOrdenada, "REPORTE AUTOMÁTICO CRON");
        }
    } catch (error) {
        console.error("❌ Error en cron de alertas:", error);
    }
}, { timezone: "America/Caracas" });

// ==========================================
// 4. INICIO DEL SERVIDOR
// ==========================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`);
});