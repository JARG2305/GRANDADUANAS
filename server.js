const express = require('express');
const cors = require('cors'); // <--- 1. Importas la dependencia aquí
const app = express();

app.use(cors()); // <--- 2. La activas para permitir conexiones externas (como Netlify)
app.use(express.json());

// ... resto de tus rutas (como /ping y /api/notificar)

// Configuración de Firebase con tu archivo JSON
const serviceAccount = require('./serviceAccountKey.json');
initializeApp({ credential: cert(serviceAccount) });
const db = getFirestore();

// Ruta que recibe las alertas desde tu interfaz web o sistemas automáticos
app.post('/api/notificar', async (req, res) => {
    try {
        const { mensaje } = req.body;
        const TELEGRAM_BOT_TOKEN = "8631519853:AAEFJVeQtj_jlbCUOnimlVXWTDeOL0qrttU";
        const TELEGRAM_CHAT_ID = "-1003976808854";

        const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                chat_id: TELEGRAM_CHAT_ID,
                text: mensaje,
                parse_mode: 'Markdown'
            })
        });
        
        const data = await response.json();
        if (data.ok) {
            res.json({ success: true, message: "Alerta enviada a Telegram con éxito." });
        } else {
            res.status(400).json({ success: false, message: data.description });
        }
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

// Ruta de mantenimiento o prueba para mantener el servidor activo
app.get('/ping', (req, res) => {
    res.send('Servidor activo 24/7 🚀');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Servidor corriendo en el puerto ${PORT}`);
});