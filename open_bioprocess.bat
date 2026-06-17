@echo off
title Bioprocess Logging System
echo Starting server and launching app...
echo --------------------------------------------------
echo *** Please keep this window open while using the application. ***
echo --------------------------------------------------
echo.

REM Open the web browser
start "" "http://localhost:5174/"

REM Run the dev server
npm run dev
