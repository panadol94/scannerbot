@echo off
setlocal enabledelayedexpansion

REM =========================
REM CONFIG (ISI SEKALI SAHAJA)
REM =========================
set SERVICE=boda8-bot
set REGION=asia-southeast1
set PROJECT=boda8-6879
set CLOUDSQL=boda8-6879:asia-southeast1:boda8-pg

set BOT_TOKEN=8479007046:AAHb6QXBIOTEfIMHzD30xwuZvIoIlA_BMIU
set WEBHOOK_SECRET=hG2cqO5dAxHI8feK7g0UD9sjNobLkuBm
set OWNER_ID=5925622731

set DATABASE_URL=postgresql+psycopg2://boda8user:Boda8Pass12345@/boda8?host=/cloudsql/boda8-6879:asia-southeast1:boda8-pg

REM =========================
REM START
REM =========================
echo.
echo ===== [1/6] Set gcloud project =====
gcloud config set project %PROJECT%
if errorlevel 1 goto :ERR

echo.
echo ===== [2/6] Deploy Cloud Run from current folder =====
REM Pastikan kau run deploy.bat dalam folder yang ada Dockerfile / source
pushd "%~dp0"

gcloud run deploy %SERVICE% ^
  --source . ^
  --region %REGION% ^
  --platform managed ^
  --allow-unauthenticated ^
  --add-cloudsql-instances %CLOUDSQL% ^
  --update-env-vars BOT_TOKEN=%BOT_TOKEN%,WEBHOOK_SECRET=%WEBHOOK_SECRET%,OWNER_ID=%OWNER_ID%,DATABASE_URL=%DATABASE_URL%
if errorlevel 1 goto :ERR

echo.
echo ===== [3/6] Get Service URL =====
for /f "usebackq delims=" %%U in (gcloud run services describe %SERVICE% --region %REGION% --format="value(status.url)") do set SVC_URL=%%U
if "%SVC_URL%"=="" goto :ERR

echo Service URL = %SVC_URL%

echo.
echo ===== [4/6] Update PUBLIC_BASE_URL (keep others) =====
gcloud run services update %SERVICE% --region %REGION% ^
  --update-env-vars PUBLIC_BASE_URL=%SVC_URL%
if errorlevel 1 goto :ERR

echo.
echo ===== [5/6] Reset Telegram webhook (drop pending updates) =====
curl.exe -sS "https://api.telegram.org/bot%BOT_TOKEN%/deleteWebhook" -d "drop_pending_updates=true"
echo.

echo.
echo ===== [6/6] Set Telegram webhook to /telegram =====
curl.exe -sS "https://api.telegram.org/bot%BOT_TOKEN%/setWebhook" ^
  -d "url=%SVC_URL%/telegram" ^
  -d "secret_token=%WEBHOOK_SECRET%" ^
  -d "drop_pending_updates=true"
echo.

echo.
echo ===== DONE =====
echo Test webhook info:
curl.exe -sS "https://api.telegram.org/bot%BOT_TOKEN%/getWebhookInfo"
echo.
popd
pause
exit /b 0

:ERR
echo.
echo !!! DEPLOY FAILED !!!
echo - Check output above.
echo - Common cause: not logged in, wrong folder, build error, or missing Dockerfile.
popd
pause
exit /b 1