@ECHO off

SET AddinDirectory=..\LanguageSetter\bin\Debug\net472

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister %AddinDirectory%\LanguageSetter.dll
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister %AddinDirectory%\PrismTaskPanes.Host.dll