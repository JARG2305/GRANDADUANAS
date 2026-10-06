const express = require('express');
const cors = require('cors');
const app = express();

// Configuración de CORS para permitir peticiones desde Netlify y entorno local
const corsOptions = {
    origin: ['https://grandaduanas.netlify.app', 'http://localhost:3000', 'http://127.0.0.1:3000'],
    methods: ['GET', 'POST'],
    allowedHeaders: ['Content-Type']
};

app.use(cors(corsOptions));
app.use(express.json());

const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN || "8631519853:AAEFJVeQtj_jlbCUOnimlVXWTDeOL0qrttU";
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID || "-1003976808854";

// Ruta raíz o de estado para UptimeRobot (devuelve texto plano o JSON en lugar de HTML)
app.get('/', (req, res) => {
    res.status(200).json({ status: "online", message: "Servidor de Logistatus Pro activo en la nube" });
});

// Ruta de notificaciones
app.post('/api/notificar', async (req, res) => {
    try {
        const { mensaje } = req.body;
        // Lógica para enviar a Telegram...
        return res.status(200).json({ success: true, message: 'Alerta enviada a Telegram correctamente.' });
    } catch (error) {
        return res.status(500).json({ success: false, message: error.message });
    }
});

        const textoMensaje = `🚨 *LOGISTATUS PRO - ALERTA*\n\n${mensaje}`;
        const urlTelegram = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
        
        const response = await fetch(urlTelegram, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                chat_id: TELEGRAM_CHAT_ID,
                text: textoMensaje,
                parse_mode: 'Markdown'
            })
        });

        const data = await response.json();

        if (!data.ok) {
            return res.status(400).json({ success: false, message: data.description || "Error al conectar con Telegram." });
        }

        return res.status(200).json({ success: true, message: 'Alerta enviada a Telegram correctamente.' });
    } catch (error) {
        console.error('❌ Error interno en el servidor:', error);
        return res.status(500).json({ success: false, message: error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Servidor corriendo en el puerto ${PORT}`);
});