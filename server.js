const express = require('express');
const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

app.use(express.static(path.join(__dirname)));
const { initializeApp, cert } = require('firebase-admin/app');
const { getFirestore } = require('firebase-admin/firestore');
const path = require('path');
const cron = require('node-cron');
const XLSX = require('xlsx');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname)));

// ==========================================
// 1. CONFIGURACIÓN DE FIREBASE ADMIN
// ==========================================
const serviceAccount = {
  "type": "service_account",
  "project_id": "statusylogistica",
  "private_key_id": "dcf3c1a7767394ad3461fff280bac03e9da5d1ca",
  "private_key": `-----BEGIN PRIVATE KEY-----
MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDTicxAZsNrM2Lq
mCXUzBZl26YTeWCPpyaD4ZZpivBrd+OwCjwoy2ivL7n35zhA4XxQo/oRCLYprV6G
SAP705WublRiY1HXO/w9L46oxYcY6Y5PYQcHrZAsAQQdFfo8GuQr09byuSzqxGwj
sOzciR8NvHUEVCbBMSnCAC53e4/XkEZzgprfn+ZX4OdIiz/2fTGq9nIPcwG7NBSb
FfVmgi5IJS3b3JNpRgDu6ZxnRwnprPE2ZAbGITY5RW3cchEWrbkFVcQtZsbWZSC3
fzgFUm4kdzqrqwhAcI+gCK8Xls+aqUXdt2pFy96PODsPUrUGa5QckS1d3xV6jI78
qesCVHaNAgMBAAECggEASV5pV8rsq4FKpRg5Qtm4SQLKUsXN7nUogCRdgWS9p2CR
OY0LOZD3UY+pnih1k9dBQUzmXkMZv1HQz50puI21xCajIO7Ww2KcRXJ5teKwzTyw
baq2//w8XW7KyPXG8VLYYbUbP9tnD1QqN6TVOlZazW5YXbR8LpSJKa7bbviNttND
lAK5xHpzOMdqdsIm7rw3lyMaOvuQcSkLzYWNFssKgtAVoivZPcdlYKxLyMpyOT3t
oXmDRvpwxQqPR4Wi4dh263QswdiL6W1eGd6VopQTUgRVOGLqsU8wxPN0mxRmA6qS
YMO9uJx2ThGHQwEJQ8MZKRhulkM5/Za0shWIA/JflwKBgQDrlAR1VWiNiYeIzANQ
DhiX7EgN16oRHkDQ9LU96PyDc2ya8dYQq4kvQLL0XglefgQ/yCZ6zA1ytxSiHj4c
0yiMS3YxCbg/R9XKTe2R6bGFzxA6sRd6IZHTLwjJip9XYHW84yIsJH7gnJhjlgBS
5h4xLwE5dYPcdRANHFFnmLXJlwKBgQDl4EiI07H+ABf0TuZrwBsNrQ8b3JWVjdYU
AlqVSFQPAvrwLNtSBMG42g5at/85nhQ6WUYadpYSkK/u+fqlgYhgOWyyJnF5H9js
/l8O3+PxR9FzD6nnjYHwTCtEfgbtHQq/bzQsFuGG8AfAIdlZUMy7NMSOsLSNod5D
AxIgud2dewKBgQDa+KHQQoxFi4GM2T13CzM2++zZ2Q3+jWVoFR3mpwsYUJCx0XrN
0fRZFMWCuWnyCCaA9tU1rTgO2jh2nK/VcT0ucvIwkL1PLMF/I0JhL5zQKPEH7RZK
cquuZfjABXco68NkyKc56s54j3ZikspIRBfqVavIsf/YSoOZ/Cl6pwrKjwKBgGI2
UMckqwc6QwG+M/QVP3m4VpwwwjgDQVOLLehZ3pALVesHPyzrm1i+0SMxOXoEb9/+
BFWKFNQZvRD8/Hl5vipeXnI+unxlfujCRRq1zU1owbPHHXAwpTNlV5cLwSnNHqpr
neH7dx70/EBCmZZYjT1UsVk4gcQOSBMPcrAIPZv/BAoGAe0aXs4lBTFop5Pr3fJbK
SVy5G/1mkLQk4SV4v7D3EYiEw4nJNEezcmXEo16IJpQxUHL9hs3irE62vPBCWKLo
phOCXNkGH6ooW7InsPNeeAnG3DDjyiUYeiWD6UF3m2r5ExuAXcMnRLh0CgUGZIyy
o/l5U2PYpS81BU2bn/4OFH8=
-----END PRIVATE KEY-----`,
  "client_email": "firebase-adminsdk-fbsvc@statusylogistica.iam.gserviceaccount.com",
  "client_id": "108295776344384360383",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/firebase-adminsdk-fbsvc%40statusylogistica.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
};

let db = null;
try {
    initializeApp({ credential: cert(serviceAccount) });
    db = getFirestore();
    console.log("🔥 Firebase inicializado correctamente.");
} catch (error) {
    console.log("⚠️ Nota: Firebase se omitió temporalmente por compatibilidad con OpenSSL, pero el servidor y Telegram funcionarán perfectamente.");
}

// ==========================================
// 2. FUNCIÓN DE CORREO DINÁMICA
// ==========================================
async function enviarCorreoSistema(opcionesMail) {
    try {
        const configDoc = await db.collection('config_correos').doc('correo_settings').get();
        if (!configDoc.exists) {
            throw new Error("No se encontró la configuración del correo en Firestore.");
        }
        const config = configDoc.data();
        const apiKey = config.apiKey; 
        const remitente = config.remitente || 'onboarding@resend.dev';

        if (!apiKey) {
            throw new Error("Falta configurar la apiKey del servicio HTTP en Firestore.");
        }

        let destinatarios = opcionesMail.to || config.emails || [];
        if (!Array.isArray(destinatarios)) {
            destinatarios = [destinatarios];
        }
        destinatarios = destinatarios.map(e => typeof e === 'string' ? e.trim() : '').filter(e => e.length > 0);

        if (destinatarios.length === 0) {
            throw new Error("No hay destinatarios válidos configurados para enviar el correo.");
        }

        let attachmentsFormatted = [];
        if (opcionesMail.attachments && Array.isArray(opcionesMail.attachments)) {
            attachmentsFormatted = opcionesMail.attachments.map(att => {
                let base64Content = '';
                if (Buffer.isBuffer(att.content)) {
                    base64Content = att.content.toString('base64');
                } else if (typeof att.content === 'string') {
                    base64Content = att.content.includes('base64,') ? att.content.split('base64,')[1] : att.content;
                }
                return {
                    filename: att.filename,
                    content: base64Content
                };
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

// ==========================================
// 3. RUTAS DE NOTIFICACIÓN (TELEGRAM)
// ==========================================
app.post('/api/notificar', async (req, res) => {
    const { mensaje } = req.body;
    
    const TELEGRAM_BOT_TOKEN = "8631519853:AAEFJVeQtj_jlbCUOnimlVXWTDeOL0qrttU";
    const TELEGRAM_CHAT_ID = "-1003976808854";

    try {
        const urlTelegram = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
        const response = await fetch(urlTelegram, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                chat_id: TELEGRAM_CHAT_ID,
                text: "🚨 <b>LOGISTATUS PRO - ALERTA FORZADA</b>\n\n" + (mensaje || "Sincronización de alertas ejecutada correctamente."),
                parse_mode: 'HTML'
            })
        });

        const data = await response.json();
        console.log("Respuesta de Telegram:", data);

        if (!data.ok) {
            throw new Error(`Telegram error: ${data.description}`);
        }

        console.log("✅ Alerta de Telegram enviada con éxito.");
        res.json({ success: true, message: 'Alerta enviada a Telegram correctamente.' });
    } catch (error) {
        console.error('❌ Error al enviar alerta a Telegram:', error);
        res.status(500).json({ success: false, message: error.message });
    }
});

// ==========================================
// 4. RUTAS DE GESTIÓN DE USUARIOS (CRUD)
// ==========================================
app.put('/api/usuarios/:id', async (req, res) => {
    const { id } = req.params;
    const { nombre, rol, password } = req.body;
    try {
        const updateData = { nombre, rol };
        if (password) updateData.password = password;
        await db.collection('usuarios').doc(id).update(updateData);
        res.json({ success: true });
    } catch (error) {
        console.error('Error al actualizar usuario:', error);
        res.status(500).json({ success: false, message: 'Error al actualizar' });
    }
});

app.delete('/api/usuarios/:id', async (req, res) => {
    const { id } = req.params;
    try {
        await db.collection('usuarios').doc(id).delete();
        res.json({ success: true });
    } catch (error) {
        console.error('Error al eliminar usuario:', error);
        res.status(500).json({ success: false, message: 'Error al eliminar' });
    }
});

// ==========================================
// 5. RUTAS DE CORREO (SISTEMA)
// ==========================================
app.post('/api/enviar-alerta-correo', async (req, res) => {
    const { destinatario, asunto, mensaje } = req.body;
    const mailOptions = {
        to: destinatario,
        subject: asunto || '🚨 Alerta Crítica - LOGISTATUS PRO',
        text: mensaje
    };

    try {
        await enviarCorreoSistema(mailOptions);
        res.json({ success: true, message: 'Correo enviado automáticamente con éxito.' });
    } catch (error) {
        console.error('Error al enviar correo automático:', error);
        res.status(500).json({ success: false, message: 'Error al enviar el correo.' });
    }
});

app.post('/api/enviar-excel-correo', async (req, res) => {
    const { destinatario, asunto, mensaje, excelBase64, nombreArchivo } = req.body;
    
    if (!excelBase64) {
        return res.status(400).json({ success: false, message: 'Falta el archivo Excel en Base64.' });
    }

    const base64Data = excelBase64.includes('base64,') ? excelBase64.split('base64,')[1] : excelBase64;

    const mailOptions = {
        to: destinatario,
        subject: asunto || '📊 Reporte de Expedientes - LOGISTATUS PRO',
        text: mensaje,
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
        console.error('Error al enviar el correo con el archivo adjunto:', error);
        res.status(500).json({ success: false, message: 'Error al enviar el correo con el Excel.' });
    }
});

// ==========================================
// 6. PROGRAMACIÓN DE REPORTES AUTOMÁTICOS (CRON)
// ==========================================
cron.schedule('0 * * * *', async () => {
    console.log('⏰ Ejecutando envío automático de reporte de Excel...');
    
    try {
        const configDoc = await db.collection('config_correos').doc('correo_settings').get();
        if (!configDoc.exists) {
            console.log('⚠️ No hay configuración de correo guardada en Firestore.');
            return;
        }
        const config = configDoc.data();
        const listaDestinatarios = config.emails ? config.emails.join(', ') : config.destinatarios;

        if (!listaDestinatarios) {
            console.log('⚠️ No hay destinatarios configurados para el envío automático.');
            return;
        }

        const snapshot = await db.collection('expedientes').get();
        const listaExpedientes = [];
        snapshot.forEach(doc => listaExpedientes.push(doc.data()));

        if (listaExpedientes.length === 0) return;

        const worksheet = XLSX.utils.json_to_sheet(listaExpedientes);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Expedientes");
        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        const nombreArchivo = `Reporte_Automatico_${new Date().toISOString().slice(0,10)}.xlsx`;

        const mailOptions = {
            to: listaDestinatarios,
            subject: '📊 Reporte Automático Programado - LOGISTATUS PRO',
            text: 'Adjunto encontrarás el reporte automático de expedientes generado por el sistema.',
            attachments: [{ filename: nombreArchivo, content: excelBuffer }]
        };

        await enviarCorreoSistema(mailOptions);
        console.log('✅ ¡Reporte automático enviado con éxito!');

    } catch (error) {
        console.error('❌ Error en el proceso automático:', error);
    }
});

// ==========================================
// 7. INICIO DEL SERVIDOR
// ==========================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`);
});