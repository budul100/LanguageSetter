@ECHO off

SET AddinDirectory=..\..\Main\bin\Debug\net472

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister /tlb "%AddinDirectory%\LanguageSetter.dll"