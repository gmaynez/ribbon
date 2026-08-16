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
CloseApplicationsFilter=excel.exe,winword.exe,powerpnt.exe,outlook.exe,Ribbon.Broker.exe
SetupMutex=Ribbon.Setup
SetupLogging=yes
ChangesEnvironment=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Types]
Name: "full"; Description: "Full installation"
Name: "custom"; Description: "Custom installation"; Flags: iscustom

[Components]
; The order here fixes the WizardForm.ComponentsList indices used in [Code].
; PostComponentIndex below must be updated if this order changes.
Name: "broker"; Description: "Ribbon Broker (required)"; Types: full custom; Flags: fixed
Name: "grid"; Description: "Ribbon Grid for Excel"; Types: full custom
Name: "quill"; Description: "Ribbon Quill for Word"; Types: full custom
Name: "deck"; Description: "Ribbon Deck for PowerPoint"; Types: full custom
Name: "post"; Description: "Ribbon Post for classic Outlook"; Types: full custom

[Files]
Source: "{#SourceDir}\Broker\*"; DestDir: "{app}\Broker"; Components: broker; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceDir}\Grid\*"; DestDir: "{app}\Grid"; Components: grid; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceDir}\Quill\*"; DestDir: "{app}\Quill"; Components: quill; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceDir}\Deck\*"; DestDir: "{app}\Deck"; Components: deck; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceDir}\Post\*"; DestDir: "{app}\Post"; Components: post; Flags: ignoreversion recursesubdirs createallsubdirs

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\Grid"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Ribbon Grid"; Components: grid; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\Grid"; ValueType: string; ValueName: "Description"; ValueData: "Ribbon Grid for Excel"; Components: grid
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\Grid"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Components: grid
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\Grid"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:VstoManifest|Grid/Grid.vsto}"; Components: grid

Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\Quill"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Ribbon Quill"; Components: quill; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\Quill"; ValueType: string; ValueName: "Description"; ValueData: "Ribbon Quill for Word"; Components: quill
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\Quill"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Components: quill
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\Quill"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:VstoManifest|Quill/Quill.vsto}"; Components: quill

Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Deck"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Ribbon Deck"; Components: deck; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Deck"; ValueType: string; ValueName: "Description"; ValueData: "Ribbon Deck for PowerPoint"; Components: deck
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Deck"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Components: deck
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\Deck"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:VstoManifest|Deck/Deck.vsto}"; Components: deck

Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\Post"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Ribbon Post"; Components: post; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\Post"; ValueType: string; ValueName: "Description"; ValueData: "Ribbon Post for classic Outlook"; Components: post
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\Post"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Components: post
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\Post"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:VstoManifest|Post/Post.vsto}"; Components: post

[UninstallRun]
Filename: "{sys}\taskkill.exe"; Parameters: "/IM Ribbon.Broker.exe /F"; Flags: runhidden; RunOnceId: "StopRibbonBroker"

[Code]
const
  // Index of the "post" component in WizardForm.ComponentsList; must match
  // the [Components] section order: broker=0, grid=1, quill=2, deck=3, post=4.
  PostComponentIndex = 4;

var
  ClassicOutlookExe: String;
  ClassicOutlookUsable: Boolean;
  PostAdjusted: Boolean;

function VstoManifest(Param: String): String;
var
  Path: String;
begin
  Path := ExpandConstant('{app}') + '\' + Param;
  StringChangeEx(Path, '\', '/', True);
  StringChangeEx(Path, ' ', '%20', True);
  Result := 'file:///' + Path + '|vstolocal';
end;

// Strips surrounding quotes and command-line parameters from an App Paths
// default value such as '"C:\...\Office16\OUTLOOK.EXE" /safe".
function ExePathFromRawValue(const Raw: String): String;
var
  Path: String;
  Space: Integer;
begin
  Result := '';
  Path := Trim(Raw);
  if Path = '' then Exit;
  if Pos('"', Path) = 1 then
  begin
    Delete(Path, 1, 1);
    Space := Pos('"', Path);
    if Space > 0 then
      Delete(Path, Space, MaxInt);
  end
  else
  begin
    Space := Pos(' ', Path);
    if Space > 0 then
      Delete(Path, Space, MaxInt);
  end;
  Result := Path;
end;

// Finds classic OUTLOOK.EXE through App Paths in both registry views and both
// hives. New Outlook repoints these keys at olk.exe / newoutlook.exe, so the
// resolved leaf name must still be outlook.exe for classic to count.
function FindClassicOutlookExe(): String;
var
  RootKeys: array[0..1] of Integer;
  Views: array[0..1] of String;
  RootIndex: Integer;
  ViewIndex: Integer;
  Raw: String;
  Candidate: String;
begin
  Result := '';
  RootKeys[0] := HKLM;
  RootKeys[1] := HKCU;
  Views[0] := 'SOFTWARE';
  Views[1] := 'SOFTWARE\WOW6432Node';
  for RootIndex := 0 to 1 do
  begin
    for ViewIndex := 0 to 1 do
    begin
      if RegQueryStringValue(
           RootKeys[RootIndex],
           Views[ViewIndex] + '\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe',
           '', Raw) then
      begin
        Candidate := ExePathFromRawValue(Raw);
        if (Candidate <> '')
           and (LowerCase(ExtractFileName(Candidate)) = 'outlook.exe')
           and FileExists(Candidate) then
        begin
          Result := Candidate;
          Exit;
        end;
      end;
    end;
  end;
end;

// New Outlook is a WebView-based client that cannot load VSTO add-ins. When
// the user has switched over (or an administrator forces it), classic
// OUTLOOK.EXE may still exist on disk but is no longer the active client.
function IsNewOutlookChosen(): Boolean;
var
  Value: Cardinal;
begin
  Result :=
    (RegQueryDWordValue(HKCU, 'Software\Microsoft\Office\16.0\Outlook\Preferences', 'NewOutlook', Value) and (Value = 1)) or
    (RegQueryDWordValue(HKLM, 'Software\Policies\Microsoft\Office\16.0\Outlook\Preferences', 'NewOutlook', Value) and (Value = 1));
end;

function HasOfficeHost: Boolean;
begin
  Result :=
    RegKeyExists(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe') or
    RegKeyExists(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe') or
    RegKeyExists(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe') or
    RegKeyExists(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe');
end;

// The VSTO add-ins run on .NET Framework 4.8 (included with Windows 10 since
// the May 2019 Update). Release 528040 is 4.8 RTM; 4.8.1 reports higher.
function HasNetFx48(): Boolean;
var
  Release: Cardinal;
begin
  Result :=
    (RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release) and (Release >= 528040)) or
    (RegQueryDWordValue(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release) and (Release >= 528040));
end;

// VSTO add-ins are loaded by the Visual Studio 2010 Tools for Office runtime,
// which is not part of Windows and must be installed separately.
function HasVstoRuntime(): Boolean;
begin
  Result :=
    RegKeyExists(HKLM, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4') or
    RegKeyExists(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4');
end;

function InitializeSetup(): Boolean;
var
  Missing: String;
begin
  ClassicOutlookExe := FindClassicOutlookExe();
  ClassicOutlookUsable := (ClassicOutlookExe <> '') and (not IsNewOutlookChosen());
  PostAdjusted := False;

  Result := True;

  Missing := '';
  if not HasNetFx48() then
    Missing := Missing + '  - Microsoft .NET Framework 4.8' + #13#10;
  if not HasVstoRuntime() then
    Missing := Missing + '  - Microsoft Visual Studio 2010 Tools for Office runtime' + #13#10;
  if Missing <> '' then
  begin
    if MsgBox(
         'This computer is missing components that Ribbon needs to run its Office add-ins:' + #13#10#13#10 +
         Missing + #13#10 +
         'The Ribbon broker itself does not need the .NET 10 runtime because it installs self-contained.' + #13#10#13#10 +
         'You can continue installing, but the add-ins will not load in Office until the missing components are installed.',
         mbInformation, MB_OKCANCEL) = IDCANCEL then
      Result := False;
    if not Result then Exit;
  end;

  if not HasOfficeHost() then
  begin
    if MsgBox(
         'Ribbon did not find Excel, Word, PowerPoint, or Outlook on this computer.'#13#10#13#10 +
         'You can install now and open an Office application later.',
         mbInformation, MB_OKCANCEL) = IDCANCEL then
      Result := False;
  end;
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if (CurPageID = wpSelectComponents) and (not PostAdjusted) then
  begin
    PostAdjusted := True;
    if not ClassicOutlookUsable then
    begin
      WizardForm.ComponentsList.CheckItem(PostComponentIndex, coUncheck);
      if ClassicOutlookExe <> '' then
        MsgBox(
          'This computer appears to be using the new Outlook.'#13#10#13#10 +
          'The new Outlook cannot load Office add-ins, so Ribbon Post was left unselected. ' +
          'It only works with classic Outlook; you can install it later if you switch back.',
          mbInformation, MB_OK)
      else
        MsgBox(
          'Ribbon did not find classic Outlook on this computer.'#13#10#13#10 +
          'Ribbon Post was left unselected. It requires the classic desktop Outlook client.',
          mbInformation, MB_OK);
    end;
  end;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if (CurPageID = wpSelectComponents)
     and WizardForm.ComponentsList.Checked[PostComponentIndex]
     and (not ClassicOutlookUsable) then
  begin
    if MsgBox(
         'Ribbon could not confirm that classic Outlook is the active mail client on this computer.'#13#10#13#10 +
         'The new Outlook cannot load Office add-ins, so Ribbon Post would not appear there.'#13#10#13#10 +
         'Install Ribbon Post anyway?',
         mbConfirmation, MB_YESNO) = IDNO then
      WizardForm.ComponentsList.CheckItem(PostComponentIndex, coUncheck);
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
