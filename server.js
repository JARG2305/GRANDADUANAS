require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { Resend } = require('resend');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));

// --- CONFIGURACIÓN DE FIRESTORE VÍA API REST (SIN DEPENDENCIAS CONFLICTIVAS) ---
let serviceAccount = null;
try {
    const serviceAccountPath = path.join(__dirname, 'sjti-70001-firebase-adminsdk-fbsvc-67d2fe1b35.json');
    serviceAccount = JSON.parse(fs.readFileSync(serviceAccountPath, 'utf8'));
    console.log("✅ Credenciales de Firestore cargadas correctamente.");
} catch (error) {
    console.warn("⚠️ No se pudo cargar el archivo de credenciales de Firestore:", error.message);
}

// Función auxiliar para obtener un token de acceso OAuth2 usando la llave privada de Firebase Service Account
async function getAccessToken() {
    if (!serviceAccount) return null;
    const crypto = require('crypto');
    const now = Math.floor(Date.now() / 1000);
    const expiry = now + 3600;

    const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
    const payload = Buffer.from(JSON.stringify({
        iss: serviceAccount.client_email,
        sub: serviceAccount.client_email,
        aud: serviceAccount.token_uri,
        iat: now,
        exp: expiry,
        scope: 'https://www.googleapis.com/auth/datastore'
    })).toString('base64url');

    const signatureInput = `${header}.${payload}`;
    const sign = crypto.createSign('RSA-SHA256');
    sign.update(signatureInput);
    sign.end();
    const signature = sign.sign(serviceAccount.private_key, 'base64url');

    const jwt = `${signatureInput}.${signature}`;

    const params = new URLSearchParams();
    params.append('grant_type', 'urn:ietf:params:oauth:grant-type:jwt-bearer');
    params.append('assertion', jwt);

    try {
        const response = await fetch(serviceAccount.token_uri, {
            method: 'POST',
            body: params
        });
        const data = await response.json();
        return data.access_token;
    } catch (e) {
        console.error("Error obteniendo token de acceso para Firestore:", e);
        return null;
    }
}

// Inicializar Resend de forma segura
let resend = null;
if (process.env.RESEND_API_KEY) {
    resend = new Resend(process.env.RESEND_API_KEY);
    console.log("✅ Resend inicializado correctamente.");
}

// --- RUTAS DE CONFIGURACIÓN ---
app.get('/api/config', (req, res) => {
    res.json({
        success: true,
        apiKey: "AIzaSyB9joRLTEKqv04ETwiuk1sYwbOhR9XuZmU",
        authDomain: "sjti-70001.firebaseapp.com",
        projectId: serviceAccount ? serviceAccount.project_id : "sjti-70001",
        storageBucket: "sjti-70001.appspot.com",
        messagingSenderId: "70001",
        appId: "1:70001:web:local",
        cloudConnected: serviceAccount !== null
    });
});

// --- ENDPOINT DE AUTENTICACIÓN (LOGIN) ---
app.post('/api/login', async (req, res) => {
    try {
        const { usuario, password } = req.body;
        if (!usuario || !password) {
            return res.status(400).json({ success: false, error: "Usuario y contraseña requeridos" });
        }

        let usuarioEncontrado = null;
        const usuarioTrim = usuario.trim().toLowerCase();

        // Intentar buscar en Firestore por REST API
        const token = await getAccessToken();
        if (token && serviceAccount) {
            const url = `https://firestore.googleapis.com/v1/projects/${serviceAccount.project_id}/databases/(default)/documents/usuarios`;
            const firestoreRes = await fetch(url, {
                headers: { 'Authorization': `Bearer ${token}` }
            });
            const firestoreData = await firestoreRes.json();
            
            if (firestoreData.documents) {
                firestoreData.documents.forEach(doc => {
                    const fields = doc.fields;
                    const dbUser = fields.usuario && fields.usuario.stringValue;
                    const dbPass = fields.password && fields.password.stringValue;
                    const dbNombre = fields.nombre && fields.nombre.stringValue;
                    const dbRol = fields.rol && fields.rol.stringValue;

                    if (dbUser && dbUser.trim().toLowerCase() === usuarioTrim && dbPass === password) {
                        usuarioEncontrado = {
                            id: doc.name.split('/').pop(),
                            usuario: dbUser,
                            nombre: dbNombre,
                            rol: dbRol
                        };
                    }
                });
            }
        }

        if (!usuarioEncontrado && usuarioTrim === 'admin' && password === 'admin') {
            usuarioEncontrado = { usuario: 'admin', nombre: 'Administrador Principal', rol: 'Administrador' };
        }

        if (usuarioEncontrado) {
            return res.json({ success: true, user: usuarioEncontrado, message: "Acceso concedido" });
        } else {
            return res.status(401).json({ success: false, error: "Usuario o contraseña incorrectos" });
        }
    } catch (error) {
        console.error("Error crítico en el login:", error);
        res.status(500).json({ success: false, error: "Error en el servidor al autenticar" });
    }
});

// --- ENDPOINTS DE GESTIÓN DE USUARIOS (FIRESTORE VÍA REST) ---
app.get('/api/usuarios', async (req, res) => {
    try {
        let usuarios = [];
        const token = await getAccessToken();
        
        if (token && serviceAccount) {
            const url = `https://firestore.googleapis.com/v1/projects/${serviceAccount.project_id}/databases/(default)/documents/usuarios`;
            const firestoreRes = await fetch(url, {
                headers: { 'Authorization': `Bearer ${token}` }
            });
            const firestoreData = await firestoreRes.json();

            if (firestoreData.documents) {
                firestoreData.documents.forEach(doc => {
                    const fields = doc.fields || {};
                    usuarios.push({
                        id: doc.name.split('/').pop(),
                        usuario: fields.usuario ? (fields.usuario.stringValue || '') : '',
                        nombre: fields.nombre ? (fields.nombre.stringValue || '') : '',
                        rol: fields.rol ? (fields.rol.stringValue || '') : '',
                        password: fields.password ? (fields.password.stringValue || '') : ''
                    });
                });
            }
        }
        
        if (usuarios.length === 0) {
            usuarios.push({ 
                id: 'local-admin-1', 
                usuario: 'admin', 
                nombre: 'Administrador Principal', 
                rol: 'Administrador' 
            });
        }
        
        return res.json(usuarios);
    } catch (error) {
        console.error("Error al obtener usuarios:", error);
        return res.json([{ 
            id: 'local-admin-1', 
            usuario: 'admin', 
            nombre: 'Administrador Principal', 
            rol: 'Administrador' 
        }]);
    }
});

// Iniciar servidor
app.listen(PORT, () => {
    console.log(`🚀 LOGISTATUS PRO corriendo en el puerto ${PORT} sin dependencias conflictivas.`);
});