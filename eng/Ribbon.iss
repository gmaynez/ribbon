#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif
#ifndef SourceDir
  #define SourceDir "..\artifacts\release\installer"
#endif

#define AppName "Ribbon"
#define AppPublisher "Ribbon"
#define AppURL "https://github.com/gmaynez/ribbon"

[Setup]
AppId={{C3F8A91D-2E47-4B6A-9D15-8F0E6C1A4B72}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}/issues
AppUpdatesURL={#AppURL}/releases
AppCopyright=Copyright (C) Ribbon contributors
LicenseFile=..\LICENSE
DefaultDirName={localappdata}\Ribbon
DefaultGroupName=Ribbon
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
SetupArchitecture=x64
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
TimeStampsInUTC=yes
OutputBaseFilename=Ribbon-Setup-v{#AppVersion}
SolidCompression=yes
Compression=lzma2
WizardStyle=modern
UninstallDisplayName={#AppName}
UninstallDisplayIcon={app}\Broker\Ribbon.Broker.exe
UsePreviousAppDir=yes
DirExistsWarning=no
CloseApplications=yes
CloseApplicationsFilter=excel.exe,winword.exe,powerpnt.exe,Ribbon.Broker.exe
SetupMutex=Ribbon.Setup
SetupLogging=yes
ChangesEnvironment=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "{#SourceDir}\Grid\*"; DestDir: "{app}\Grid"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceDir}\Quill\*"; DestDir: "{app}\Quill"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceDir}\Deck\*"; DestDir: "{app}\Deck"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceDir}\Broker\*"; DestDir: "{app}\Broker"; Flags: ignoreversion recursesubdirs createallsubdirs

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\Grid"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Ribbon Grid"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\Grid"; ValueType: string; ValueName: "Description"; ValueData: "Ribbon Grid for Excel"
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\Grid"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\Grid"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:VstoManifest|Grid/Grid.vsto}"

Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\Quill"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Ribbon Quill"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\Quill"; ValueType: string; ValueName: "Description"; ValueData: "Ribbon Quill for Word"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\Quill"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\Quill"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:VstoManifest|Quill/Quill.vsto}"

Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Deck"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Ribbon Deck"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Deck"; ValueType: string; ValueName: "Description"; ValueData: "Ribbon Deck for PowerPoint"
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Deck"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Deck"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:VstoManifest|Deck/Deck.vsto}"

[UninstallRun]
Filename: "{sys}\taskkill.exe"; Parameters: "/IM Ribbon.Broker.exe /F"; Flags: runhidden; RunOnceId: "StopRibbonBroker"

[Code]
function VstoManifest(Param: String): String;
var
  Path: String;
begin
  Path := ExpandConstant('{app}') + '\' + Param;
  StringChangeEx(Path, '\', '/', True);
  StringChangeEx(Path, ' ', '%20', True);
  Result := 'file:///' + Path + '|vstolocal';
end;

function HasOfficeHost: Boolean;
begin
  Result :=
    RegKeyExists(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe') or
    RegKeyExists(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe') or
    RegKeyExists(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe');
end;

function InitializeSetup(): Boolean;
begin
  Result := True;
  if not HasOfficeHost() then
  begin
    if MsgBox(
         'Ribbon did not find Excel, Word, or PowerPoint on this computer.'#13#10#13#10 +
         'You can install now and open an Office application later.',
         mbInformation, MB_OKCANCEL) = IDCANCEL then
      Result := False;
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ResultCode: Integer;
begin
  NeedsRestart := False;
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/IM Ribbon.Broker.exe /F', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Result := '';
end;
