@echo off
title Excel to CSV Column Extractor

:: Find the JAR file in target directory
set JAR_FILE=target\excel-to-csv-1.0.0.jar

:: Check if JAR exists
if not exist "%JAR_FILE%" (
    echo ERROR: JAR file not found at %JAR_FILE%
    echo Please run 'mvn package' first to build the application.
    pause
    exit /b 1
)

:: Launch the application
java -jar "%JAR_FILE%" %*

:: If launched with no arguments (GUI mode), don't pause
if "%~1"=="" exit /b 0

:: For CLI mode, pause to show results
pause
