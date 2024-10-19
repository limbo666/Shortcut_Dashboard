@echo off
setlocal enabledelayedexpansion

REM Set variables
set "wallpaperFileName=wallpaper1.jpg"
set "sourceFolder=%~dp0Wallpapers"
set "destinationFolder=%USERPROFILE%\Pictures\Wallpapers"
set "wallpaperPath=%destinationFolder%\%wallpaperFileName%"

REM Check for admin rights
NET SESSION >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Administrator privileges required. Relaunching with elevated permissions...
    powershell -Command "Start-Process '%~dpnx0' -Verb RunAs"
    exit /b
)

REM Copy Wallpapers folder content to Pictures\Wallpapers
if exist "%sourceFolder%" (
    if not exist "%destinationFolder%" mkdir "%destinationFolder%"
    robocopy "%sourceFolder%" "%destinationFolder%" /E /IS /IT
    if %ERRORLEVEL% LEQ 1 (
        echo Files copied successfully.
    ) else (
        echo Warning: Some files may not have copied. Check the output above for details.
    )
) else (
    echo Source Wallpapers folder not found
    exit /b
)

REM Check if wallpaper file exists
if not exist "%wallpaperPath%" (
    echo Wallpaper file not found: %wallpaperPath%
    echo Please check if the file exists and try again.
    timeout /t 5
    exit /b
)

REM Set the wallpaper
reg add "HKCU\Control Panel\Desktop" /v Wallpaper /t REG_SZ /d "%wallpaperPath%" /f

REM Update User Profile
rundll32.exe user32.dll,UpdatePerUserSystemParameters ,1 ,True

REM Force refresh using multiple methods
RUNDLL32.EXE user32.dll, UpdatePerUserSystemParameters
RUNDLL32.EXE user32.dll, UpdatePerUserSystemParameters 1, True

REM Additional refresh attempts
for /L %%i in (1,1,3) do (
    RUNDLL32.EXE user32.dll, UpdatePerUserSystemParameters
    timeout /t 1 >nul
)

echo Wallpaper has been set and desktop refreshed.
echo Wallpaper path: %wallpaperPath%

REM Pause for 3 seconds and then close
echo.
echo Script completed. Closing in 3 seconds...
timeout /t 3 >nul
exit