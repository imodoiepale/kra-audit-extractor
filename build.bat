@echo off
echo ========================================
echo KRA POST PORTUM TOOL - Build Script
echo ========================================
echo.

echo [1/3] Checking dependencies...
call npm install
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)

echo.
echo [2/3] Building desktop application...
call npm run build
if %errorlevel% neq 0 (
    echo ERROR: Build failed
    pause
    exit /b 1
)

echo.
echo [3/3] Build complete!
echo.
echo ========================================
echo Your installers are ready in the dist/ folder:
echo.
echo - KRA POST PORTUM TOOL Setup 1.0.0.exe (Installer)
echo - KRA POST PORTUM TOOL 1.0.0.exe (Portable)
echo ========================================
echo.
pause
