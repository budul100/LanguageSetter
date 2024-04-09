#include <idp.iss>

#define ProgramName "LanguageSetter"
#define ProgramVersion "1.1.2"
#define ProgramPublisher "budul"

#define setupPath "..\..\Main\bin\Release"

#define PTPHostName "PrismTaskPanes.Host"
#define DownloadUrlNet6 "https://download.visualstudio.microsoft.com/download/pr/9d6b6b34-44b5-4cf4-b924-79a00deb9795/2f17c30bdf42b6a8950a8552438cf8c1/windowsdesktop-runtime-6.0.6-win-x64.exe"

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
[CustomMessages]
IDP_DownloadFailed=Download of .NET 6 failed. .NET 6 Desktop runtime is required to run VidCoder.
IDP_RetryCancel=Click 'Retry' to try downloading the files again, or click 'Cancel' to terminate setup.
InstallingDotNetRuntime=Installing .NET 6 Desktop Runtime. This might take a few minutes...
DotNetRuntimeFailedToLaunch=Failed to launch .NET Runtime Installer with error "%1". Please fix the error then run this installer again.
DotNetRuntimeFailed1602=.NET Runtime installation was cancelled. This installation can continue, but be aware that this application may not run unless the .NET Runtime installation is completed successfully.
DotNetRuntimeFailed1603=A fatal error occurred while installing the .NET Runtime. Please fix the error, then run the installer again.
DotNetRuntimeFailed5100=Your computer does not meet the requirements of the .NET Runtime. Please consult the documentation.
DotNetRuntimeFailedOther=The .NET Runtime installer exited with an unexpected status code "%1". Please review any other messages shown by the installer to determine whether the installation completed successfully, and abort this installation and fix the problem if it did not.

[Code]
var
  requiresRestart: boolean;

function CompareVersion(V1, V2: string): Integer;
var
  P, N1, N2: Integer;
begin
  Result := 0;
  while (Result = 0) and ((V1 <> '') or (V2 <> '')) do
  begin
    P := Pos('.', V1);
    if P > 0 then
    begin
      N1 := StrToInt(Copy(V1, 1, P - 1));
      Delete(V1, 1, P);
    end
      else
    if V1 <> '' then
    begin
      N1 := StrToInt(V1);
      V1 := '';
    end
      else
    begin
      N1 := 0;
    end;
    P := Pos('.', V2);
    if P > 0 then
    begin
      N2 := StrToInt(Copy(V2, 1, P - 1));
      Delete(V2, 1, P);
    end
      else
    if V2 <> '' then
    begin
      N2 := StrToInt(V2);
      V2 := '';
    end
      else
    begin
      N2 := 0;
    end;
    if N1 < N2 then Result := -1
      else
    if N1 > N2 then Result := 1;
  end;
end;

function NetRuntimeIsMissing(): Boolean;
var
  runtimes: TArrayOfString;
  registryKey: string;
  I: Integer;
  meetsMinimumVersion: Boolean;
  meetsMaximumVersion: Boolean;
  minimumVersion: string;
  maximumExclusiveVersion: string;
begin
  Result := True;

  minimumVersion := '6.0.0';
  maximumExclusiveVersion := '6.1.0';
  registryKey := 'SOFTWARE\WOW6432Node\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App';
  if RegGetValueNames(HKLM, registryKey, runtimes) then
  begin
    for I := 0 to GetArrayLength(runtimes)-1 do
    begin
      meetsMinimumVersion := not (CompareVersion(runtimes[I], minimumVersion) = -1);
      meetsMaximumVersion := CompareVersion(runtimes[I], maximumExclusiveVersion) = -1;
      if meetsMinimumVersion and meetsMaximumVersion then
      begin
        Log(Format('[.NET] Selecting %s', [runtimes[I]]));
        Result := False;
          Exit;
      end;
    end;
  end;
end;

procedure InitializeWizard;
begin
  if NetRuntimeIsMissing() then
  begin
    idpAddFile(ExpandConstant('{#DownloadUrlNet6}'), ExpandConstant('{tmp}\NetRuntimeInstaller.exe'));
    idpDownloadAfter(wpReady);
  end;
end;

function InstallDotNetRuntime(): String;
var
  StatusText: string;
  ResultCode: Integer;
begin
  StatusText := WizardForm.StatusLabel.Caption;
  WizardForm.StatusLabel.Caption := CustomMessage('InstallingDotNetRuntime');
  WizardForm.ProgressGauge.Style := npbstMarquee;
  try
    if not Exec(ExpandConstant('{tmp}\NetRuntimeInstaller.exe'), '/passive /norestart /showrmui /showfinalerror', '', SW_SHOW, ewWaitUntilTerminated, ResultCode) then
    begin
      Result := FmtMessage(CustomMessage('DotNetRuntimeFailedToLaunch'), [SysErrorMessage(resultCode)]);
    end
    else
    begin
      // See https://msdn.microsoft.com/en-us/library/ee942965(v=vs.110).aspx#return_codes
      case resultCode of
        0: begin
          // Successful
        end;
        1602 : begin
          MsgBox(CustomMessage('DotNetRuntimeFailed1602'), mbInformation, MB_OK);
        end;
        1603: begin
          Result := CustomMessage('DotNetRuntimeFailed1603');
        end;
        1641: begin
          requiresRestart := True;
        end;
        3010: begin
          requiresRestart := True;
        end;
        5100: begin
          Result := CustomMessage('DotNetRuntimeFailed5100');
        end;
        else begin
          MsgBox(FmtMessage(CustomMessage('DotNetRuntimeFailedOther'), [IntToStr(resultCode)]), mbError, MB_OK);
        end;
      end;
    end;
  finally
    WizardForm.StatusLabel.Caption := StatusText;
    WizardForm.ProgressGauge.Style := npbstNormal;

    DeleteFile(ExpandConstant('{tmp}\NetRuntimeInstaller.exe'));
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  // 'NeedsRestart' only has an effect if we return a non-empty string, thus aborting the installation.
  // If the installers indicate that they want a restart, this should be done at the end of installation.
  // Therefore we set the global 'restartRequired' if a restart is needed, and return this from NeedRestart()

  if NetRuntimeIsMissing() then
  begin
    Result := InstallDotNetRuntime();
  end;
end;

function NeedRestart(): Boolean;
begin
  Result := requiresRestart;
end;

