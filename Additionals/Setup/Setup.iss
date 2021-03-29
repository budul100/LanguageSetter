#include ".\Base.iss"

#define ProgramName "LanguageSetter"
#define ProgramVersion "1.0.3"
#define ProgramPublisher "budul"

#define PrismTaskPanesHost "PrismTaskPanes.Host"
#define setupPath "..\..\Main\bin\Release\net472"

#define UseDotNet47

#define UseNetCoreCheck
#ifdef UseNetCoreCheck
  #define UseNetCore31
  #define UseNetCore31Desktop
#endif

[Setup]
AppId={{9E46AC4E-2659-464E-997B-1F00D1741939}
AppName={#ProgramName}
AppVersion={#ProgramVersion}
AppVerName={#ProgramName} {#ProgramVersion}
AppPublisher={#ProgramPublisher}
DefaultDirName={pf32}\{#ProgramName}
DefaultGroupName={#ProgramName}
OutputDir=..\..
OutputBaseFilename={#ProgramName}_Setup
Compression=lzma
SolidCompression=yes
ChangesAssociations=True
AlwaysUsePersonalGroup=True
VersionInfoVersion={#ProgramVersion}
VersionInfoCompany={#ProgramPublisher}
AllowUNCPath=False
AlwaysShowGroupOnReadyPage=True
AlwaysShowDirOnReadyPage=True
DisableWelcomePage=no
UninstallDisplayIcon={uninstallexe}
ChangesEnvironment=True

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"
Name: "de"; MessagesFile: "compiler:Languages\German.isl"

[Files]
Source: "{#setupPath}\{#ProgramName}.dll"; DestDir: "{app}"; Flags: ignoreversion 
Source: "{#setupPath}\{#PrismTaskPanesHost}.dll"; DestDir: "{commoncf32}\{#PrismTaskPanesHost}"; Flags: ignoreversion sharedfile 
Source: "{#setupPath}\*.dll"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{cm:UninstallProgram,{#ProgramName}}"; Filename: "{uninstallexe}"

[Run]
Filename: "{dotnet4032}\RegAsm.exe"; Parameters: """{app}\{#ProgramName}.dll"""; WorkingDir: "{app}"; Flags: runhidden; StatusMsg: "Register addin libraries"
Filename: "{dotnet4032}\RegAsm.exe"; Parameters: "/codebase ""{commoncf32}\{#PrismTaskPanesHost}\{#PrismTaskPanesHost}.dll"""; WorkingDir: "{commoncf32}\{#PrismTaskPanesHost}"; Flags: runhidden; StatusMsg: "Register addin libraries"
Filename: "{code:GetPowerpointPath}"; Flags: nowait postinstall skipifsilent unchecked; Description: "{cm:LaunchPowerPoint,PowerPoint}"

[UninstallRun]
Filename: "{dotnet4032}\RegAsm.exe"; Parameters: "/unregister ""{app}\{#ProgramName}.dll"""; Flags: runhidden; StatusMsg: "Unregister addin libraries"
Filename: "{dotnet4032}\RegAsm.exe"; Parameters: "/unregister ""{commoncf32}\{#PrismTaskPanesHost}\{#PrismTaskPanesHost}.dll"""; Flags: runhidden; StatusMsg: "Unregister addin libraries"

[CustomMessages]
de.LaunchPowerPoint=Starte Microsoft PowerPoint nach der Installation
en.LaunchPowerPoint=Start Microsoft PowerPoint after finishing installation

[ThirdParty]
UseRelativePaths=True

[Code]
function GetPowerpointPath(dummy: string): string;
begin
  RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe', '', Result);
  if Result = '' then Result := 'powerpnt.exe';
end;
