@echo off
:: MUFAP Pipeline Launcher
:: Automatically loads variables from config.env and executes the pipeline.

cd /d "%~dp0"

:: Read config.env file line by line and set variables
if exist config.env (
    for /f "usebackq tokens=1,* delims==" %%A in ("config.env") do (
        echo %%A | findstr /r "^#" >nul || (
            set "%%A=%%B"
        )
    )
) else if exist .env (
    for /f "usebackq tokens=1,* delims==" %%A in (".env") do (
        echo %%A | findstr /r "^#" >nul || (
            set "%%A=%%B"
        )
    )
)

:: Execute the script
python main.py
pause