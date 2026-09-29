const express = require('express');
const cors = require('cors');
const cron = require('node-cron');

const app = express();
app.use(cors());
app.use(express.json());

const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN || "8631519853:AAEFJVeQtj_jlbCUOnimlVXWTDeOL0qrttU";
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID || "-5559170176";
const FIREBASE_PROJECT_ID = "statusylogistica";

// Función auxiliar para enviar mensajes a Telegram
async function enviarAlertaTelegram(mensaje) {
    try {
        const urlTelegram = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
        const response = await fetch(urlTelegram, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                chat_id: TELEGRAM_CHAT_ID,
                text: mensaje,
                parse_mode: 'Markdown'
            })
        });
        const data = await response.json();
        if (!response.ok) {
            console.error("Error al enviar a Telegram:", data);
        }
    } catch (err) {
        console.error("Fallo de red en Telegram:", err);
    }
}

// 1. Endpoint manual que recibe peticiones del botón "Forzar Alertas" del Frontend
app.post('/api/notificar', async (req, res) => {
    try {
        const { mensaje } = req.body;
        if (!mensaje) {
            return res.status(400).json({ error: "Falta el mensaje" });
        }
        await enviarAlertaTelegram(mensaje);
        res.json({ success: true });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.get('/', (req, res) => {
    res.send('LOGISTATUS PRO - Backend de Alertas Activo 🚀');
});

// 2. Cron Job Automatizado (Corre cada hora en el minuto 0)
cron.schedule('0 * * * *', async () => {
    const ahora = new Date();
    const horaActualStr = ahora.toLocaleString('en-US', { timeZone: 'America/Caracas', hour: 'numeric', hour12: false });
    const horaNum = parseInt(horaActualStr, 10);

    // Franja horaria permitida: de 9:00 AM a 8:00 PM
    if (horaNum < 9 || horaNum > 20) return;

    let debeEnviar = horaNum <= 18 ? ((horaNum - 9) % 3 === 0) : true;
    if (!debeEnviar) return;

    console.log(`⏰ [CRON AUTOMÁTICO] Verificando expedientes en Firebase para las ${horaNum}:00 hrs...`);

    try {
        // Consultar expedientes directamente desde Firestore vía REST API
        const urlFirestore = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/expedientes`;
        const firestoreRes = await fetch(urlFirestore);
        const firestoreData = await firestoreRes.json();

        if (!firestoreData.documents) {
            console.log("ℹ️ No hay expedientes registrados en la base de datos.");
            return;
        }

        const hoy = new Date();
        hoy.setHours(0, 0, 0, 0);
        let alertasEnviadasCount = 0;

        for (const doc of firestoreData.documents) {
            const fields = doc.fields || {};
            const expName = fields.num_expediente?.stringValue || "S/N";
            const clienteInfo = fields.cliente?.stringValue || "S/N";
            const vencimientoDai = fields.vencimiento_dai?.stringValue || "";
            const fechaLlegada = fields.fecha_llegada?.stringValue || "";
            const despacho = fields.despacho?.stringValue || "";

            if (despacho) continue; // Si ya está despachado, se ignora para alertas operativas

            // Evaluar Vencimiento DAI (<= 3 días)
            if (vencimientoDai) {
                let fechaLimpia = vencimientoDai;
                if (fechaLimpia.includes('/')) {
                    const partes = fechaLimpia.split('/');
                    if (partes.length === 3) fechaLimpia = `${partes[2]}-${partes[1]}-${partes[0]}`;
                }
                const fechaDai = new Date(fechaLimpia + 'T00:00:00');
                if (!isNaN(fechaDai)) {
                    const diffDias = Math.round((fechaDai - hoy) / (1000 * 60 * 60 * 24));
                    if (diffDias <= 3) {
                        let diasTexto = diffDias < 0 ? "¡VENCIDO!" : (diffDias === 0 ? "Vence HOY" : `Faltan ${diffDias} días`);
                        const mensajeDai = `⚠️ *DAI PRÓXIMO A VENCER*\n\nExp: *${expName}*\nCliente: *${clienteInfo}*\nEl DAI vence el *${vencimientoDai}* (${diasTexto}).`;
                        await enviarAlertaTelegram(mensajeDai);
                        alertasEnviadasCount++;
                    }
                }
            }

            // Evaluar Arribo / Llegada
            if (fechaLlegada) {
                const mensajeArribo = `🚢 **ARRIBO DE CARGA REGISTRADO**\n\nExp: *${expName}*\nCliente: *${clienteInfo}*\nFecha de llegada prevista/registrada: *${fechaLlegada}*.`;
                // Puedes agregar aquí una validación por fecha si deseas filtrar por cuenta regresiva exacta
            }
        }

        console.log(`✅ Rutina automática completada. Se enviaron ${alertasEnviadasCount} alerta(s) de forma autónoma.`);
    } catch (error) {
        console.error("❌ Error en la tarea programada del servidor:", error);
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Servidor corriendo en el puerto ${PORT}`);
});