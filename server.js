const express = require('express');
const cors = require('cors'); // <--- 1. Importar cors
const { initializeApp, cert } = require('firebase-admin/app');
const { getFirestore } = require('firebase-admin/firestore');
const path = require('path');
const cron = require('node-cron');
const XLSX = require('xlsx');

const app = express();
app.use(cors()); // <--- 2. Activar cors para que Netlify pueda comunicarse con Render
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname)));

// ==========================================
// 1. CONFIGURACIÓN DE FIREBASE ADMIN (VÍA JSON)
// ==========================================
const serviceAccount = require('./serviceAccountKey.json');

let db = null;
try {
    initializeApp({ credential: cert(serviceAccount) });
    db = getFirestore();
    console.log("🔥 Firebase inicializado correctamente.");
} catch (error) {
    console.error("❌ Error al inicializar Firebase:", error);
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Servidor corriendo en http://localhost:${PORT}`);
});