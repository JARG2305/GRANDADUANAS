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
const cron = require('node-cron');
const nodemailer = require('nodemailer');

// Configuración del servicio de correo (puedes usar Gmail o cualquier SMTP corporativo)
const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: process.env.CORREO_USER || 'statusepga@gmail.com', // Tu correo
        pass: process.env.CORREO_PASS || 'kambpqjmczwfmtyj' // Contraseña de aplicación de Gmail
    }
});

// Función para enviar el reporte consolidado por correo
async function enviarCorreoAutomatico(datosAlertas) {
    if (!datosAlertas || datosAlertas.length === 0) return;

    // Ordenar las alertas por ETA en orden ascendente
    datosAlertas.sort((a, b) => new Date(a.fechaObj) - new Date(b.fechaObj));

    let htmlContenido = `
        <div style="font-family: Arial, sans-serif; color: #333; padding: 20px;">
            <h2 style="color: #0056b3;">📊 Logistatus Pro - Reporte Consolidado de Alertas</h2>
            <p>A continuación se detallan las alertas operativas ordenadas por fecha de llegada (ETA) más cercana:</p>
            <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
                <thead>
                    <tr style="background-color: #0056b3; color: white;">
                        <th style="padding: 10px; border: 1px solid #ddd;">Expediente</th>
                        <th style="padding: 10px; border: 1px solid #ddd;">Cliente</th>
                        <th style="padding: 10px; border: 1px solid #ddd;">Tipo / Modo</th>
                        <th style="padding: 10px; border: 1px solid #ddd;">Detalle / ETA</th>
                    </tr>
                </thead>
                <tbody>
    `;

    datosAlertas.forEach(alerta => {
        htmlContenido += `
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd; text-align: center;"><b>${alerta.expediente}</b></td>
                <td style="padding: 8px; border: 1px solid #ddd;">${alerta.cliente}</td>
                <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${alerta.modo}</td>
                <td style="padding: 8px; border: 1px solid #ddd;">${alerta.detalle}</td>
            </tr>
        `;
    });

    htmlContenido += `
                </tbody>
            </table>
            <p style="margin-top: 20px; font-size: 12px; color: #777;">Este es un reporte automático generado por Logistatus Pro en la nube.</p>
        </div>
    `;

    const mailOptions = {
        from: '"Logistatus Pro" <statusepga@gmail.com>',
        to: 'importacionesepga@gmail.com',
        subject: '🚀 Reporte Automático de Alertas Operativas - Logistatus Pro',
        html: htmlContenido
    };

    try {
        await transporter.sendMail(mailOptions);
        console.log("📧 Correo automático enviado exitosamente.");
    } catch (error) {
        console.error("❌ Error al enviar el correo automático:", error);
    }
}

// ==========================================
// PROGRAMACIÓN AUTOMÁTICA CON NODE-CRON
// ==========================================
// Se ejecuta todos los días a las 10:00 a.m. y a las 6:00 p.m. (Hora del servidor)
cron.schedule('0 10,18 * * *', async () => {
    console.log("⏰ Ejecutando tarea programada: Verificando y enviando alertas...");
    
    // Aquí puedes conectar tu lógica para extraer los registros activos de Firebase 
    // o recibirlos desde tu frontend, ordenarlos por ETA ascendente y pasarlos a la función:
    // const registrosPendientes = await obtenerRegistrosDeBaseDe Datos();
    // await enviarCorreoAutomatico(registrosPendientes);
    
}, {
    scheduled: true,
    timezone: "America/Caracas" // Ajusta a tu zona horaria local
});