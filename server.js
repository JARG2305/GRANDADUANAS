const express = require('express');
const path = require('path');
const XLSX = require('xlsx');
const cron = require('node-cron');

const app = express();

// ==========================================
// 0. CONFIGURACIÓN DE MIDDLEWARES
// ==========================================
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use(express.static(path.join(__dirname)));

// Credenciales del Bot de Telegram
const TELEGRAM_BOT_TOKEN = "8631519853:AAEFJVeQtj_jlbCUOnimlVXWTDeOL0qrttU";
const TELEGRAM_CHAT_ID = "-1003976808854";

// ==========================================
// 1. FUNCIONES DE SERVICIO (CORREO Y TELEGRAM)
// ==========================================
async function enviarCorreoSistema(opcionesMail) {
    try {
        const apiKey = "re_fvWGsDmy_CxGHq2taybxAxuQ9Eb2fadPQ"; 
        const remitente = "onboarding@resend.dev";

        if (!apiKey || apiKey.includes("TU_API_KEY")) {
            throw new Error("Falta configurar tu apiKey de Resend en el servidor.");
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
            console.log("✅ Alerta automática enviada a Telegram desde la nube.");
        }
    } catch (err) {
        console.error("❌ Error de red al conectar con Telegram desde el servidor:", err);
    }
}

// ==========================================
// 2. RUTAS DE CORREO (MANUALES)
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

// ==========================================
// 3. TAREAS PROGRAMADAS EN SEGUNDO PLANO (CRON)
// ==========================================

// Tarea 1: Envío de correos automáticos a las 10:00 AM y 6:00 PM (18:00)
cron.schedule('0 10,18 * * *', async () => {
    console.log("⏰ [CRON] Ejecutando envío programado de reporte a las 10 AM / 6 PM...");
    try {
        const destinatariosAutomaticos = ["importacionesepga@gmail.com", "hnoguera@gmail.com", "finanzascepga@gmail.com"];
        const mailOptionsAuto = {
            to: destinatariosAutomaticos,
            subject: "Reporte Maestro Unificado de Expedientes - LOGISTATUS PRO",
            text: "Este es el envío automático programado del reporte consolidado de operaciones.",
            html: "<p>Este es el envío automático programado del reporte consolidado de operaciones de <strong>Logistatus Pro</strong>.</p>"
        };
        await enviarCorreoSistema(mailOptionsAuto);
        console.log("✅ [CRON] Correo automático programado enviado con éxito.");
    } catch (error) {
        console.error("❌ [CRON] Error al enviar el correo automático programado:", error);
    }
});

// Tarea 2: Alertas de Telegram a las 9, 12, 15, 18, 19 y 20 horas
cron.schedule('0 9,12,15,18,19,20 * * *', async () => {
    const horaActual = new Date().getHours();
    console.log(`⏰ [CRON] Ejecutando verificación de alertas de Telegram (${horaActual}:00 hrs)...`);
    try {
        await enviarAlertaTelegramServidor(`⏰ **LOGISTATUS PRO - REVISIÓN AUTOMÁTICA EN NUBE**\n\nVerificación programada de las ${horaActual}:00 horas ejecutada en segundo plano.`);
    } catch (error) {
        console.error("❌ [CRON] Error al procesar las alertas automáticas de Telegram:", error);
    }
});

// ==========================================
// 4. INICIO DEL SERVIDOR
// ==========================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`);
});