#include <idp.iss>

#define ProgramName "LanguageSetter"
#define ProgramVersion "1.2.0"
#define ProgramPublisher "budul"

#define setupPath "..\..\Main\bin\Release"
#define PTPHostName "PrismTaskPanes.Host"

; requires netcorecheck.exe and netcorecheck_x64.exe (see CodeDependencies.iss)
#define public Dependency_Path_NetCoreCheck "dependencies\"
#include "CodeDependencies.iss"

[Setup]
AppId={{9E46AC4E-2659-464E-997B-1F00D1741939}
AppName={#ProgramName}
AppVersion={#ProgramVersion}
AppVerName={#ProgramName} {#ProgramVersion}
AppPublisher={#ProgramPublisher}
DefaultDirName={pf}\{#ProgramName}
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
ArchitecturesInstallIn64BitMode=x64
DisableProgramGroupPage=yes

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"
Name: "de"; MessagesFile: "compiler:Languages\German.isl"

[Files]
Source: "{#setupPath}\*.*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs
Source: "{#setupPath}\{#PTPHostName}.*"; DestDir: "{commoncf64}\{#PTPHostName}"; Flags: ignoreversion sharedfile
Source: "{#setupPath}\{#PTPHostName}.comhost.dll"; DestDir: "{commoncf64}\{#PTPHostName}"; Flags: ignoreversion regserver sharedfile
Source: "{#setupPath}\{#ProgramName}.*"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#setupPath}\{#ProgramName}.comhost.dll"; DestDir: "{app}"; Flags: ignoreversion regserver

[Icons]
Name: "{group}\{cm:UninstallProgram,{#ProgramName}}"; Filename: "{uninstallexe}"

[ThirdParty]
UseRelativePaths=True

[Code]

function InitializeSetup: Boolean;
begin

#ifdef Dependency_Path_NetCoreCheck
  Dependency_AddDotNet80Desktop;
#endif

  Result := True;

end;
