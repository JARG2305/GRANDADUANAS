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
// 0.4 MOTOR CENTRAL DE COMANDOS Y ALERTAS
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
                    detalleDaiTexto = `\n⚠️ *FALTA REGISTRAR DAI*`;
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

if (textoLower.includes('razón social:') || textoLower.includes('razon social:') || textoLower.includes('rif:')) {
        try {
            const razonSocialMatch = textoMensaje.match(/(?:Razón Social|Razon Social):\s*([^\n]+)/i);
            const rifMatch = textoMensaje.match(/RIF:\s*([^\n]+)/i);
            const direccionMatch = textoMensaje.match(/Dirección:\s*([^\n]+)/i);
            const telefonoMatch = textoMensaje.match(/Teléfono:\s*([^\n]+)/i);
            const correoMatch = textoMensaje.match(/Correo:\s*([^\n]+)/i);

            const nuevoCliente = {
                razonSocial: razonSocialMatch ? razonSocialMatch[1].trim() : "S/N",
                rif: rifMatch ? rifMatch[1].trim() : "S/N",
                direccion: direccionMatch ? direccionMatch[1].trim() : "",
                telefono: telefonoMatch ? telefonoMatch[1].trim() : "",
                correo: correoMatch ? correoMatch[1].trim() : ""
            };

            const guardadoExitoso = await guardarClienteFirestore(nuevoCliente);
            if (guardadoExitoso) {
                return `✅ *¡CLIENTE REGISTRADO CON ÉXITO EN LA NUBE!* 🚀\n\n` +
                       `🏢 *Razón Social:* ${nuevoCliente.razonSocial}\n` +
                       `📄 *RIF:* ${nuevoCliente.rif}\n` +
                       `📍 *Dirección:* ${nuevoCliente.direccion}\n` +
                       `📞 *Teléfono:* ${nuevoCliente.telefono}\n` +
                       `📧 *Correo:* ${nuevoCliente.correo}`;
            } else {
                return `❌ Hubo un error al guardar el cliente en la base de datos de Firestore.`;
            }
        } catch (e) {
            console.error("Error procesando nuevo cliente:", e);
        }
    }

    const palabrasAyuda = ['comandos', 'ayuda', 'menu', 'menú', 'sistema'];
    const coincidenciasAyuda = palabrasEnMensaje.filter(p => palabrasAyuda.includes(p)).length;

    if (coincidenciasAyuda >= 2 || textoLower === 'ayuda' || textoLower === 'comandos') {
        return `🤖 *GUÍA DE COMANDOS - LOGISTATUS PRO* 🤖\n\n` +
               `🚨 *1. Alertas Activas*\n• *"alertas activas"* | *"alerta sistema"*\n\n` +
               `📊 *2. Resumen*\n• *"resumen exp"*\n\n` +
               `✅ *3. Indicadores (Sí / No)*\n` +
               `• *"Retirado bl [Exp o BL]"* | *"Valorado [Exp o BL]"* | *"Factura recibida [Exp o BL]"* | *"Prevaloración enviada [Exp o BL]"*\n\n` +
               `🎨 *4. Canal Aduanero*\n` +
               `• *"Canal [Exp o BL] verde"*, *"amarillo"*, *"rojo"*\n\n` +
               `👤 *5. Funcionario*\n` +
               `• *"Funcionario exp 26-030 [Nombre]"*\n\n` +
               `✏️ *6. Fechas (Evaluando Secuencia Estricta)*\n` +
               `• *"Llegada [Exp o BL] al..."* ➔ *Requiere ETA*\n` +
               `• *"Acta / Recibido acta [Exp o BL] al..."* ➔ *Requiere Llegada*\n` +
               `• *"Transmitido / Transmisión [Exp o BL] al..."* ➔ *Requiere Acta*\n` +
               `• *"Reconocido [Exp o BL] al..."* ➔ *Requiere Transmisión*\n` +
               `• *"Validado [Exp o BL] al..."* ➔ *Requiere Reconocimiento*\n` +
               `• *"Despachado [Exp o BL] al..."* ➔ *Requiere Validación*`;
    }

    const palabrasAlertas = ['alertas', 'alerta', 'activas', 'exp', 'expedientes', 'sistema', 'urgentes', 'vencimientos'];
    const coincidenciasAlertas = palabrasEnMensaje.filter(p => palabrasAlertas.includes(p)).length;

    if (coincidenciasAlertas >= 2) {
        const listaFinalOrdenada = await generarListaAlertasSistema();
        if (listaFinalOrdenada.length === 0) return "✅ *¡EXCELENTE!* No hay alertas activas.";
        let mensajesEnumerados = listaFinalOrdenada.map((item, index) => `${index + 1}️⃣ ${item}`);
        return `🚨 *ALERTAS ACTIVAS (ORDENADAS POR ETA/VENCIMIENTO)* 🚨\n\n` + mensajesEnumerados.join("\n\n");
    }

    const palabrasResumen = ['resumen', 'exp', 'expedientes', 'activos', 'conteo', 'totales', 'proceso'];
    const coincidenciasResumen = palabrasEnMensaje.filter(p => palabrasResumen.includes(p)).length;

    if (coincidenciasResumen >= 2) {
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) return "✅ No hay registros in Firestore.";

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
    // 4. DETECCIÓN Y ACTUALIZACIÓN DIRECTA (CON EVALUACIÓN DE SECUENCIA ESTRICTA)
    // =========================================================================
    let campoAModificar = "";
    let nombreCampoLegible = "";
    let esIndicadorSiNo = false;
    let esCanal = false;
    let esFuncionario = false;
    let valorCanal = "";
    let esBorrar = textoLower.includes('borrar') || textoLower.includes('eliminar') || textoLower.includes('quitar');

    const esNegativo = textoLower.startsWith('no ') || textoLower.includes(' no ');

    if (textoLower.includes('retirado bl') || textoLower.includes('retiro bl') || textoLower.includes('recibido bl') || textoLower.includes('doc transporte')) {
        campoAModificar = "recibido_documento_transporte";
        nombreCampoLegible = "Recibido Doc. Transporte";
        esIndicadorSiNo = true;
    } else if (textoLower.includes('prevaloración enviada') || textoLower.includes('prevaloracion enviada')) {
        campoAModificar = "prevaloracion_enviada";
        nombreCampoLegible = "Prevaloración Enviada";
        esIndicadorSiNo = true;
    } else if (textoLower.includes('valorado')) {
        campoAModificar = "documento_valorado_en_sistema";
        nombreCampoLegible = "Doc. Valorado en Sistema";
        esIndicadorSiNo = true;
    } else if (textoLower.includes('factura recibida') || textoLower.includes('recibida factura')) {
        campoAModificar = "factura_recibida";
        nombreCampoLegible = "Factura Recibida";
        esIndicadorSiNo = true;
    } else if (textoLower.includes('canal')) {
        campoAModificar = "canal";
        nombreCampoLegible = "Canal Aduanero";
        esCanal = true;
        if (textoLower.includes('verde')) valorCanal = "VERDE";
        else if (textoLower.includes('amarillo')) valorCanal = "AMARILLO";
        else if (textoLower.includes('rojo')) valorCanal = "ROJO";
    } else if (textoLower.includes('funcionario')) {
        campoAModificar = "funcionario";
        nombreCampoLegible = "Funcionario";
        esFuncionario = true;
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
        campoAModificar = "recibida_acta_recepcion";
        nombreCampoLegible = "Recibida Acta de Recepción";
    }

    if (campoAModificar && (textoLower.includes('actualizar') || textoLower.includes('actualiza') || esIndicadorSiNo || esCanal || esFuncionario || esBorrar || textoLower.includes('transmitido') || textoLower.includes('reconocido') || textoLower.includes('validado') || textoLower.includes('despachado') || textoLower.includes('retirada') || textoLower.includes('canal') || textoLower.includes('valorado') || textoLower.includes('factura') || textoLower.includes('prevaloración') || textoLower.includes('retirado') || textoLower.includes('recibido') || textoLower.includes('no ') || textoLower.includes('eta') || textoLower.includes('llegada') || textoLower.includes('verde') || textoLower.includes('amarillo') || textoLower.includes('rojo') || textoLower.includes('funcionario') || textoLower.includes('acta'))) {
        
        let valorNuevo = "";
        if (esBorrar) {
            valorNuevo = "";
        } else if (esIndicadorSiNo) {
            valorNuevo = esNegativo ? "No" : "Sí";
        } else if (esCanal) {
            valorNuevo = valorCanal;
            if (!valorNuevo) return `❌ Debes indicar el color del canal (Verde, Amarillo o Rojo).`;
        } else if (esFuncionario) {
            let textoLimpio = textoMensaje.replace(/funcionario/gi, '').replace(/(exp|bl)\s*[\w-]+/gi, '').trim();
            valorNuevo = textoLimpio;
            if (!valorNuevo) return `❌ Debes indicar el nombre del funcionario a registrar.`;
        } else {
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

        // ⚠️ AQUÍ ESTÁ LA CORRECCIÓN: Debe evaluarse dentro de este "if"
        if (!expedienteEncontrado) {
            const terminoBuscado = textoMensaje.split(" ")[1] || "desconocido";
            return generarPlantillaClienteNuevo(terminoBuscado);
        }
        // =========================================================================
        // EVALUACIÓN ESTRICTA DE SECUENCIA (EVENTO ANTERIOR OBLIGATORIO)
        // =========================================================================
        const f = expedienteEncontrado.fields;
        const hasEta = Boolean(f.eta_la_guaira?.stringValue && f.eta_la_guaira.stringValue.trim() !== "");
        const hasLlegada = Boolean(f.fecha_llegada?.stringValue && f.fecha_llegada.stringValue.trim() !== "");
        const hasActa = Boolean(f.recibida_acta_recepcion?.stringValue && f.recibida_acta_recepcion.stringValue.trim() !== "");
        const hasTransmision = Boolean(f.fecha_transmision?.stringValue && f.fecha_transmision.stringValue.trim() !== "");
        const hasCanal = Boolean(f.canal?.stringValue && f.canal.stringValue.trim() !== "");
        const hasReconocimiento = Boolean(f.reconocimiento?.stringValue && f.reconocimiento.stringValue.trim() !== "");
        const hasValidacion = Boolean(f.validacion?.stringValue && f.validacion.stringValue.trim() !== "");

        // 1. Llegada requiere ETA
        if (campoAModificar === "fecha_llegada" && !hasEta) {
            return `⚠️ *Secuencia Requerida:* No se puede registrar la *Fecha de Llegada* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene registrado el **ETA**.`;
        }
        // 2. Acta requiere Llegada
        if (campoAModificar === "recibida_acta_recepcion" && !hasLlegada) {
            return `⚠️ *Secuencia Requerida:* No se puede registrar el *Acta de Recepción* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene registrada la **Fecha de Llegada**.`;
        }
        // 3. Transmisión requiere Acta
        if (campoAModificar === "fecha_transmision" && !hasActa) {
            return `⚠️ *Secuencia Requerida:* No se puede registrar la *Fecha de Transmisión* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene registrada el **Acta de Recepción**.`;
        }
        // 4. Canal requiere Transmisión
        if (esCanal && !hasTransmision) {
            return `⚠️ *Secuencia Requerida:* No se puede asignar el *Canal* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene registrada la **Fecha de Transmisión**.`;
        }
        // 5. Reconocimiento requiere Transmisión (y Canal recomendado)
        if (campoAModificar === "reconocimiento" && !hasTransmision) {
            return `⚠️ *Secuencia Requerida:* No se puede registrar el *Reconocimiento* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene registrada la **Fecha de Transmisión**.`;
        }
        // 6. Validación requiere Reconocimiento
        if (campoAModificar === "validacion" && !hasReconocimiento) {
            return `⚠️ *Secuencia Requerida:* No se puede registrar la *Validación* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene registrado el **Reconocimiento**.`;
        }
        // 7. Despacho requiere Validación
        if (campoAModificar === "despacho" && !hasValidacion) {
            return `⚠️ *Secuencia Requerida:* No se puede registrar el *Despacho* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene registrada la **Validación**.`;
        }
        // 8. Funcionario requiere Canal asignado
        if (esFuncionario && !hasCanal) {
            return `⚠️ *Secuencia Requerida:* No se puede asignar el *Funcionario* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene asignado un **Canal** aduanero.`;
        }

        const camposActualizados = { ...expedienteEncontrado.fields };
        camposActualizados[campoAModificar] = { stringValue: valorNuevo };

        const updateUrl = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/expedientes/${expedienteEncontrado.id}`;
        const patchBody = { fields: camposActualizados };

        const patchRes = await fetch(updateUrl, {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(patchBody)
        });

        if (patchRes.ok) {
            if (esBorrar) {
                return `🗑️ *¡CAMPO ELIMINADO CON ÉXITO!* ⚠️\n\n📂 Exp: *${expedienteEncontrado.numExpediente}*\n👤 Cliente: *${expedienteEncontrado.cliente}*\n📦 B/L: *${expedienteEncontrado.bl}*\n⚙️ Campo: *${nombreCampoLegible}* (Ha quedado vacío).`;
            } else if (esIndicadorSiNo) {
                return `✅ *¡INDICADOR ACTUALIZADO A ${valorNuevo.toUpperCase()}!* 🚀\n\n📂 Exp: *${expedienteEncontrado.numExpediente}*\n👤 Cliente: *${expedienteEncontrado.cliente}*\n📦 B/L: *${expedienteEncontrado.bl}*\n⚙️ Campo: *${nombreCampoLegible}* ➔ *${valorNuevo.toUpperCase()}*`;
            } else if (esCanal) {
                return `🎨 *¡CANAL ACTUALIZADO CON ÉXITO!* 🚀\n\n📂 Exp: *${expedienteEncontrado.numExpediente}*\n👤 Cliente: *${expedienteEncontrado.cliente}*\n📦 B/L: *${expedienteEncontrado.bl}*\n⚙️ Campo: *Canal Aduanero* ➔ *${valorNuevo.toUpperCase()}*`;
            } else if (esFuncionario) {
                return `👤 *¡FUNCIONARIO ASIGNADO CON ÉXITO!* 🚀\n\n📂 Exp: *${expedienteEncontrado.numExpediente}*\n👤 Cliente: *${expedienteEncontrado.cliente}*\n📦 B/L: *${expedienteEncontrado.bl}*\n⚙️ Funcionario: *${valorNuevo}*`;
            } else {
                return `✅ *¡FECHA ACTUALIZADA CON ÉXITO!* 🚀\n\n📂 Exp: *${expedienteEncontrado.numExpediente}*\n👤 Cliente: *${expedienteEncontrado.cliente}*\n📦 B/L: *${expedienteEncontrado.bl}*\n⚙️ Campo: *${nombreCampoLegible}*\n📅 Fecha Actualizada: *${formatearFechaDDMMYYYY(valorNuevo)}*`;
            }
        } else {
            const errBody = await patchRes.text();
            console.error("❌ Error Firestore PATCH:", errBody);
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

// Obtener todos los clientes desde la web o backend
app.get('/api/clientes', async (req, res) => {
    const clientes = await obtenerClientesFirestore();
    res.json({ success: true, clientes });
});

// Guardar un cliente nuevo desde la web
app.post('/api/clientes', async (req, res) => {
    const exito = await guardarClienteFirestore(req.body);
    res.json({ success: exito });
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

// Validación cuando un cliente no es encontrado en la base de datos
function generarPlantillaClienteNuevo(nombreBuscado) {
    return `⚠️ *CLIENTE NO ENCONTRADO EN LA BASE DE DATOS* ⚠️\n\n` +
           `El cliente o expediente "*${nombreBuscado}*" no se encuentra registrado.\n\n` +
           `Por favor, envíe los datos con el siguiente formato exacto para agregarlo:\n\n` +
           `📌 *REGISTRAR CLIENTE*\n` +
           `• *Razón Social:* [Nombre o Razón Social]\n` +
           `• *RIF:* [Ej: J-XXXXXXXX-X]\n` +
           `• *Dirección:* [Dirección fiscal completa]\n` +
           `• *Teléfono:* [Número de contacto]\n` +
           `• *Correo:* [Correo electrónico]`;
}
// ==========================================
// GESTIÓN DE CLIENTES EN FIRESTORE (NUBE)
// ==========================================

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
                rif: fields.rif?.stringValue || '',
                direccion: fields.direccion?.stringValue || '',
                telefono: fields.telefono?.stringValue || '',
                correo: fields.correo?.stringValue || ''
            };
        });
    } catch (err) {
        console.error("❌ Error obteniendo clientes de Firestore:", err);
        return [];
    }
}

async function guardarClienteFirestore(datosCliente) {
    try {
        const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/clientes`;
        const body = {
            fields: {
                razonSocial: { stringValue: datosCliente.razonSocial || '' },
                rif: { stringValue: datosCliente.rif || '' },
                direccion: { stringValue: datosCliente.direccion || '' },
                telefono: { stringValue: datosCliente.telefono || '' },
                correo: { stringValue: datosCliente.correo || '' },
                fechaCreacion: { stringValue: new Date().toISOString() }
            }
        };

        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });

        return response.ok;
    } catch (err) {
        console.error("❌ Error guardando cliente en Firestore:", err);
        return false;
    }
}
// ==========================================
// 3. TAREAS PROGRAMADAS (CRON)
// ==========================================
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
// Función que escanea automáticamente todas las variables de Render y arranca el bot de cada cliente
function iniciarTodosLosBotsAutomaticos() {
    console.log("🤖 [MULTIBOT] Escaneando variables de Render para iniciar los bots...");

    // Recorremos todo process.env buscando variables que terminen en _TOKEN o _BOT_TOKEN
    for (const [key, value] of Object.entries(process.env)) {
        if ((key.endsWith('_TOKEN') || key.endsWith('_BOT_TOKEN')) && value) {
            const tokenBot = value;
            // Extraemos el prefijo (ej: de SOTO_BOT_TOKEN o SOTO_TOKEN sacamos SOTO)
            const prefijoCliente = key.replace(/(_BOT_TOKEN|_TOKEN)$/, '');
            const chatIdEsperado = process.env[`${prefijoCliente}_CHAT_ID`];

            console.log(`🚀 [BOT DETECTADO] Activando vigilancia para el cliente: ${prefijoCliente}`);
            
            // Iniciamos un ciclo de escucha independiente para este bot
            ejecutarPollingParaBot(tokenBot, chatIdEsperado, prefijoCliente);
        }
    }
}

// Ciclo de escucha independiente por cada bot encontrado
function ejecutarPollingParaBot(tokenBot, chatIdEsperado, nombreCliente) {
    let offset = 0;

    setInterval(async () => {
        try {
            const url = `https://api.telegram.org/bot${tokenBot}/getUpdates?offset=${offset}&timeout=30`;
            const response = await fetch(url);
            const data = await response.json();

            if (data.ok && data.result.length > 0) {
                for (const update of data.result) {
                    offset = update.update_id + 1;
                    if (!update.message || !update.message.text) continue;

                    const chatIdRecibido = String(update.message.chat.id);
                    const textoMensaje = update.message.text.trim();

                    // Si se configuró un Chat ID específico para este cliente, validamos que coincida por seguridad
                    if (chatIdEsperado && chatIdRecibido !== String(chatIdEsperado)) {
                        continue; 
                    }

                    console.log(`💬 Mensaje recibido en el bot de [${nombreCliente}] (Chat: ${chatIdRecibido})`);

                    // Procesar el comando de logística
                    const respuestaProcesada = await procesarComandoLogistica(textoMensaje);

                    if (respuestaProcesada) {
                        await enviarRespuestaTelegramCliente(tokenBot, chatIdRecibido, respuestaProcesada);
                    }
                }
            }
        } catch (err) {}
    }, 4000);
}

async function enviarRespuestaTelegramCliente(token, chatIdDestino, mensajeTexto) {
    if (!chatIdDestino || !token) return;
    try {
        const url = `https://api.telegram.org/bot${token}/sendMessage`;
        await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ chat_id: chatIdDestino, text: mensajeTexto, parse_mode: 'Markdown' })
        });
    } catch (err) {
        console.error("❌ Error respondiendo en Telegram:", err);
    }
}

// Iniciar todos los bots automáticamente al encender el servidor
iniciarTodosLosBotsAutomaticos();

// ==========================================
// 4. INICIO DEL SERVIDOR
// ==========================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => { console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`); });