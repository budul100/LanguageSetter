@ECHO off

CLS
ECHO D ... Debug
ECHO R ... Release
ECHO.

CHOICE /C DR /M "Enter the evironment to be registred:"

IF ERRORLEVEL 2 GOTO Release
IF ERRORLEVEL 1 GOTO Debug

:Debug
SET AddinDir=..\..\Main\bin\Debug
GOTO Register

:Release
SET AddinDir=..\..\Main\bin\Release
GOTO Register

:Register

regsvr32 "%AddinDir%\PrismTaskPanes.Host.comhost.dll"
regsvr32 "%AddinDir%\LanguageSetter.comhost.dll"