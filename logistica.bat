@echo off
title LOGISTATUS PRO - Servidor y Escritorio
cd /d "F:\pen\prueba epga/index.html"
start cmd /k "npm run dev"
timeout /t 3 >nul
start http://localhost:3000/
exit