require('dotenv').config();
const express = require('express');
const app = express(); 
const path = require('path');
const XLSX = require('xlsx-js-style');
const cron = require('node-cron');
const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode');

// ==========================================
// 0. CONFIGURACIÓN DE MIDDLEWARES Y WHATSAPP
// ==========================================
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use(express.static(path.join(__dirname)));

const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID;
const FIREBASE_PROJECT_ID = process.env.FIREBASE_PROJECT_ID || "statusylogistica";
const FIREBASE_API_KEY = process.env.FIREBASE_API_KEY || "";

app.use((req, res, next) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    if (req.method === 'OPTIONS') {
        return res.sendStatus(200);
    }
    next();
});

// Inicialización única de WhatsApp
let whatsappClientReady = false;
let ultimoQrGenerado = "";
let estadoConexionWhatsapp = "DESCONECTADO";

const whatsappClient = new Client({
    authStrategy: new LocalAuth({ clientId: "logistatus-pro-session", dataPath: ".wwebjs_auth" }),
    puppeteer: {
        headless: true,
        executablePath: process.env.CHROME_PATH || undefined,
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-dev-shm-usage',
            '--disable-accelerated-2d-canvas',
            '--no-first-run',
            '--no-zygote',
            '--single-process',
            '--disable-gpu'
        ]
    }
});

whatsappClient.on('qr', async (qr) => {
    console.log('📱 [WHATSAPP] Nuevo código QR generado.');
    estadoConexionWhatsapp = "ESPERANDO_ESCANEO";
    try {
        ultimoQrGenerado = await qrcode.toDataURL(qr);
    } catch (err) {
        console.error("❌ Error al generar QR:", err);
    }
});

whatsappClient.on('ready', () => {
    console.log('✅ [WHATSAPP] Conectado y listo en la nube.');
    whatsappClientReady = true;
    estadoConexionWhatsapp = "CONECTADO";
    ultimoQrGenerado = "";
});

whatsappClient.on('auth_failure', (msg) => {
    console.error('❌ [WHATSAPP] Error de autenticación:', msg);
    estadoConexionWhatsapp = "ERROR_AUTENTICACION";
});

whatsappClient.on('disconnected', (reason) => {
    console.log('⚠️ [WHATSAPP] Desconectado:', reason);
    whatsappClientReady = false;
    estadoConexionWhatsapp = "DESCONECTADO";
    whatsappClient.initialize();
});

whatsappClient.initialize();