@echo off
cd /d "%~dp0"

echo Starting WhatsApp Service...
start "WhatsApp Service" cmd /k "cd /d "%~dp0whatsapp-service" && npm start"

timeout /t 4 /nobreak > nul

echo Starting Cloudflare Tunnel...
echo.
echo ============================================================
echo  Copy the https://xxxx.trycloudflare.com URL below and
echo  paste it into Streamlit Cloud secrets as WA_SERVICE_URL
echo ============================================================
echo.
cloudflared.exe tunnel --url http://localhost:3001
pause
