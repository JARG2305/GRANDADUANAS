const express = require('express');
const cors = require('cors');
const cron = require('node-cron');

const app = express();
app.use(cors());
app.use(express.json());

// Token y Chat ID de tu bot de Telegram (puedes usar variables de entorno en Render)
const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN || "TU_TOKEN_DEL_BOT";
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID || "TU_CHAT_ID";

// 1. Endpoint principal que recibe las peticiones para notificar a Telegram
app.post('/api/notificar', async (req, res) => {
    try {
        const { mensaje } = req.body;
        if (!mensaje) {
            return res.status(400).json({ error: "Falta el mensaje" });
        }

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
            return res.status(500).json({ error: "Error en Telegram", details: data });
        }

        res.json({ success: true, result: data });
    } catch (err) {
        console.error("Fallo de red en /api/notificar:", err);
        res.status(500).json({ error: err.message });
    }
});

// Endpoint básico para mantener el servicio activo (Keep-Alive)
app.get('/', (req, res) => {
    res.send('LOGISTATUS PRO - Backend de Alertas Activo 🚀');
});

// 2. Cron Job automatizado (Corre cada hora en el minuto 0)
cron.schedule('0 * * * *', async () => {
    const ahora = new Date();
    // Zona horaria ajustada (ej: Venezuela / America/Caracas)
    const horaActualStr = ahora.toLocaleString('en-US', { timeZone: 'America/Caracas', hour: 'numeric', hour12: false });
    const horaNum = parseInt(horaActualStr, 10);

    // Franja permitida: de 9:00 AM (9) a 8:00 PM (20)
    if (horaNum < 9 || horaNum > 20) {
        return; 
    }

    // Regla de frecuencia:
    // - De 9:00 AM a 6:00 PM (9, 12, 15, 18): Cada 3 horas
    // - De 7:00 PM a 8:00 PM (19, 20): Cada 1 hora
    let debeEnviar = false;
    if (horaNum <= 18) {
        debeEnviar = ((horaNum - 9) % 3 === 0);
    } else {
        debeEnviar = true; // 19 y 20
    }

    if (!debeEnviar) return;

    console.log(`⏰ [CRON AUTOMÁTICO] Ejecutando verificación para las ${horaNum}:00 hrs...`);

    try {
        // Aquí puedes realizar la consulta a Firebase (o disparar la lógica de revisión)
        // Y enviar la notificación automática a Telegram mediante fetch interno si es necesario.
        console.log("✅ Rutina de horarios evaluada con éxito en el servidor.");
    } catch (error) {
        console.error("❌ Error en la tarea programada del servidor:", error);
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Servidor corriendo en el puerto ${PORT}`);
});