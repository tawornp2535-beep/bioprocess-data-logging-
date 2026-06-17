@echo off
title Stop Bioprocess System
echo Stopping Bioprocess System servers...

REM Find and kill process on port 5000 (Express Backend)
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :5000 ^| findstr LISTENING') do (
    taskkill /f /pid %%a 2>nul
    echo Stopped backend server on port 5000.
)

REM Find and kill process on port 5174 (Vite Frontend)
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :5174 ^| findstr LISTENING') do (
    taskkill /f /pid %%a 2>nul
    echo Stopped frontend server on port 5174.
)

echo.
echo Done! Bioprocess servers have been stopped.
ping 127.0.0.1 -n 4 >nul
