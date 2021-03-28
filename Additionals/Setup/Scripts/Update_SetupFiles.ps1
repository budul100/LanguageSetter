param (
	[string] $projectPath,
	[string] $issPath,
	[string] $updaterPath
)

$scriptpath = split-path -parent $MyInvocation.MyCommand.Definition

. $scriptpath\Get_VersionCurrent.ps1

$currentVersion = Get-ProjectVersion -projectPath $projectPath

. $scriptpath\Set_FileContent.ps1

Set-FileContent -path $issPath -replace '(?<=\#define\ ProgramVersion\ )(\"[^\"]*\")' -replaceWith """$currentVersion"""
Set-FileContent -path $updaterPath -replace '(?<=<version>)[^<]+' -replaceWith "$currentVersion"

exit 0