const express = require('express');
const { initializeApp, cert } = require('firebase-admin/app');
const { getFirestore } = require('firebase-admin/firestore');
const path = require('path');
const cron = require('node-cron');
const XLSX = require('xlsx');

const app = express();

// ==========================================
// 0. CONFIGURACIÓN DE MIDDLEWARES
// ==========================================
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
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
    console.log("⚠️ Nota: Firebase se omitió temporalmente.");
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
            throw new Error("No hay destinatarios válidos configurados.");
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

// ==========================================
// 3. RUTAS DE NOTIFICACIÓN Y CORREO
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
        console.error('Error al enviar el correo:', error);
        res.status(500).json({ success: false, message: error.message || 'Error al enviar el correo.' });
    }
});

// ==========================================
// 4. INICIO DEL SERVIDOR
// ==========================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`);
});