@echo off

SET UpdateScripts=.\Additionals\Scripts
SET SetupDir=.\Additionals\Setup
SET SetupPath=%SetupDir%\Setup.iss
SET SetupScripts=%SetupDir%\Scripts

SET ProjectPath=.\Main\LanguageSetter.csproj
SET SlnPaths='%ProjectPath%'

SET OutputDir1=.
SET OutputDir2=.

echo.
echo Clean solution
echo.

SET BaseDir=%~dp0

powershell "%UpdateScripts%\_CleanFolders.ps1" -baseDir "%CD%"

RMDIR /S /Q "%APPDATA%\LanguageSetter"

echo.
echo Build solution
echo.

CHOICE /C mb /N /M "Shall the [b]uild (x.x._X_) or the [m]inor version (x._X_.0) be increased?"
SET VERSIONSELECTION=%ERRORLEVEL%
echo.

if /i "%VERSIONSELECTION%" == "1" (
	echo.
	echo Update minor version
	echo.

	powershell "%SetupScripts%\Update_VersionMinor.ps1 -projectPaths %SlnPaths%"
)

dotnet build "%ProjectPath%" --configuration Release

echo.
echo Test solution
echo.

dotnet test LanguageSetter.sln

if not %ERRORLEVEL% == 0 goto EndProcess

echo.
echo Update setup file
echo.

powershell -command "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted; exit 0"
powershell "%SetupScripts%\Update_SetupFiles.ps1 -projectPath '%ProjectPath%' -issPath '%SetupPath%'"

echo.
echo Compile setup file
echo.

"%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" "%SetupPath%"

if not %ERRORLEVEL% == 0 goto EndProcess

echo.
echo Deploy setup file
echo.

copy /y ".\LanguageSetter_Setup.exe" "%OutputDir1%"
copy /y ".\LanguageSetter_Setup.exe" "%OutputDir2%"

powershell "%SetupScripts%\Update_VersionBuild.ps1 -projectPaths %SlnPaths%"

echo.
echo.
echo COMPILING AND DEPLOYING WAS SUCCESFULL.
echo.
echo.

:EndProcess

pause

start "" https://sourceforge.net/projects/languagesetter/files/