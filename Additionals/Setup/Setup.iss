#include ".\Base.iss"

#define ProgramName "Language Setter"
#define ProgramVersion "0.1.0"
#define ProgramPublisher "budul"

#define PrismTaskPanesName "PrismTaskPanes"
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
Source: "{#setupPath}\{#AddInName}.dll"; DestDir: "{app}"; Flags: ignoreversion regtypelib
Source: "{#setupPath}\{#PrismTaskPanesName}.dll"; DestDir: "{app}"; Flags: ignoreversion sharedfile regtypelib
Source: "{#setupPath}\*.dll"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{cm:UninstallProgram,{#ProgramName}}"; Filename: "{uninstallexe}"

[CustomMessages]
de.LaunchPowerPoint=Starte Microsoft PowerPoint nach der Installation
en.LaunchPowerPoint=Start Microsoft PowerPoint after finishing installation

[ThirdParty]
UseRelativePaths=True

[Run]
Filename: "{code:GetPowerpointPath}"; Flags: nowait postinstall skipifsilent unchecked; Description: "{cm:LaunchPowerPoint,PowerPoint}"

[UninstallDelete]
Type: files; Name: "{app}\lib\{#AddInName}.tlb"
Type: files; Name: "{app}\lib\{#PrismTaskPanesName}.tlb"

[Code]
function GetPowerpointPath(dummy: string): string;
begin
  RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe', '', Result);
  if Result = '' then Result := 'powerpnt.exe';
end;

function GetUninstallString(): String;
var
  sUnInstPath: String;
  sUnInstallString: String;
begin
  sUnInstPath := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{#emit SetupSetting("AppId")}_is1');
  sUnInstallString := '';
  if not RegQueryStringValue(HKLM, sUnInstPath, 'UninstallString', sUnInstallString) then
    RegQueryStringValue(HKCU, sUnInstPath, 'UninstallString', sUnInstallString);
  Result := sUnInstallString;
end;

function IsUpgrade(): Boolean;
begin
  Result := (GetUninstallString() <> '');
end;

function UnInstallOldVersion(): Integer;
var
  sUnInstallString: String;
  iResultCode: Integer;
begin
  // Return Values:
  // 1 - uninstall string is empty
  // 2 - error executing the UnInstallString
  // 3 - successfully executed the UnInstallString

  // default return value
  Result := 0;

  // get the uninstall string of the old app
  sUnInstallString := GetUninstallString();
  if sUnInstallString <> '' then begin
    sUnInstallString := RemoveQuotes(sUnInstallString);
    if Exec(sUnInstallString, '/SILENT /NORESTART /SUPPRESSMSGBOXES','', SW_HIDE, ewWaitUntilTerminated, iResultCode) then
      Result := 3
    else
      Result := 2;
  end else
    Result := 1;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if (CurStep=ssInstall) then
  begin
    if (IsUpgrade()) then
    begin
      UnInstallOldVersion();
    end;
  end;
end;
