const express = require('express');
const cors = require('cors');
const { initializeApp, cert } = require('firebase-admin/app');
const { getFirestore } = require('firebase-admin/firestore');
const path = require('path');
const cron = require('node-cron');
const XLSX = require('xlsx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname)));

// ==========================================
// 1. CONFIGURACIÓN DE FIREBASE ADMIN (SEGURO PARA RENDER Y LOCAL)
// ==========================================
let serviceAccount;

if (process.env.FIREBASE_SERVICE_ACCOUNT) {
    try {
        serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
    } catch (error) {
        console.error("❌ Error al parsear la variable FIREBASE_SERVICE_ACCOUNT:", error);
    }
} else {
    try {
        serviceAccount = require('./serviceAccountKey.json');
    } catch (error) {
        console.warn("⚠️ No se encontró el archivo serviceAccountKey.json local.");
    }
}

let db = null;
try {
    if (serviceAccount) {
        initializeApp({ credential: cert(serviceAccount) });
        db = getFirestore();
        console.log("🔥 Firebase inicializado correctamente.");
    } else {
        console.error("❌ No se pudieron cargar las credenciales de Firebase.");
    }
} catch (error) {
    console.error("❌ Error al inicializar Firebase:", error);
}

// ==========================================
// 2. RUTA DE BIENVENIDA (RAÍZ)
// ==========================================
app.get('/', (req, res) => {
    res.send('¡Servidor de GRANDADUANAS en línea y operativo! 🚀');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Servidor corriendo en http://localhost:${PORT}`);
});