@ECHO off

SET AddinDir=..\..\Main\bin\Debug

regsvr32 /u "%AddinDir%\LanguageSetter.comhost.dll"
regsvr32 /u "%AddinDir%\PrismTaskPanes.Host.comhost.dll"

SET AddinDir=..\..\Main\bin\Release

regsvr32 /u "%AddinDir%\LanguageSetter.comhost.dll"
regsvr32 /u "%AddinDir%\PrismTaskPanes.Host.comhost.dll"