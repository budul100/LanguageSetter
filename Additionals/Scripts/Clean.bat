@echo off

RMDIR /S /Q "%APPDATA%\LanguageSetter"

SET BaseDir=%~dp0..\..

pushd %~dp0

powershell ".\_CleanFolders.ps1 -baseDir %BaseDir%"

popd