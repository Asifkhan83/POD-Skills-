@echo off
REM POD Daily Workflow - Double-click to run
REM Edit the paths below to match your environment

SET POD_FOLDER=D:\PODs
SET MANIFEST=D:\Data\manifest.xlsx
SET SKILLS_DIR=%~dp0

echo ============================================================
echo POD Daily Workflow
echo ============================================================
echo.
echo POD Folder: %POD_FOLDER%
echo Manifest:   %MANIFEST%
echo.

REM Check if manifest exists
if not exist "%MANIFEST%" (
    echo ERROR: Manifest file not found: %MANIFEST%
    echo Please edit this batch file and set the correct MANIFEST path.
    pause
    exit /b 1
)

REM Run the workflow
python "%SKILLS_DIR%daily_workflow.py" "%MANIFEST%"

echo.
echo ============================================================
echo Workflow completed. Press any key to close.
pause
