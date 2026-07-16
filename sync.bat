@echo off
title GarageOS Sync v1.0
color 0A

echo.
echo ==========================================================
echo                    GarageOS Sync v1.0
echo ==========================================================
echo.

REM ----------------------------------------------------------
REM Root folder
REM ----------------------------------------------------------

cd /d "%~dp0"

echo Project:
cd

echo.

REM ----------------------------------------------------------
REM STEP 1 - Pull latest Apps Script
REM ----------------------------------------------------------

echo [1/5] Pulling latest Apps Script...
cd appscript

call clasp pull

if errorlevel 1 (
    echo.
    echo ERROR: Apps Script synchronization failed.
    goto END
)

cd ..

echo.
echo SUCCESS
echo.

REM ----------------------------------------------------------
REM STEP 2 - Detect changes
REM ----------------------------------------------------------

echo [2/5] Checking repository...

git status --short

git diff --quiet
if %errorlevel%==0 (
    git diff --cached --quiet
    if %errorlevel%==0 (
        echo.
        echo No changes detected.
        goto END
    )
)

echo.
echo Changes detected.
echo.

REM ----------------------------------------------------------
REM STEP 3 - Commit message
REM ----------------------------------------------------------

set /p MESSAGE=Commit message (Enter = Update GarageOS): 

if "%MESSAGE%"=="" (
    set MESSAGE=Update GarageOS
)

echo.

REM ----------------------------------------------------------
REM STEP 4 - Commit
REM ----------------------------------------------------------

echo [3/5] Staging files...

git add .

echo.
echo [4/5] Creating commit...

git commit -m "%MESSAGE%"

if errorlevel 1 (
    echo.
    echo ERROR: Commit failed.
    goto END
)

echo.
echo SUCCESS
echo.

REM ----------------------------------------------------------
REM STEP 5 - Push
REM ----------------------------------------------------------

echo [5/5] Uploading to GitHub...

git push origin main

if errorlevel 1 (
    echo.
    echo ERROR: Push failed.
    goto END
)

echo.
echo ==========================================================
echo               GARAGEOS SYNCHRONIZED
echo ==========================================================

:END

echo.
pause