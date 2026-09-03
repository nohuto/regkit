#define AppId "4678f42c-c6a2-4df9-bc2a-dddbd2613045"
#define AppName "RegKit"
#define AppExeName "regkit.exe"
#define AppVersion "0.0.0.7"
#define AppPublisher "nohuto"
#define AppCopyright "(C) 2026 nohuto"
#define AppURL "https://github.com/nohuto/regkit"
#ifndef Arch
  #define Arch "x64"
#endif
#if Arch == "x86"
  #define BuildDir "..\\build32\\Release"
#else
  #define BuildDir "..\\build\\Release"
#endif

[Setup]
AppId={#AppId}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
VersionInfoVersion={#AppVersion}
VersionInfoCopyright={#AppCopyright}
DefaultDirName={code:GetDefaultDir}
DefaultGroupName=RegKit
CreateAppDir=yes
DisableDirPage=no
DisableProgramGroupPage=yes
LicenseFile=..\LICENSE
SetupIconFile=..\assets\icons\regkit.ico
UninstallDisplayIcon={app}\{#AppExeName}
Compression=lzma2
SolidCompression=yes
#if Arch == "x64"
ArchitecturesAllowed=x64os
ArchitecturesInstallIn64BitMode=x64os
#endif
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog
WizardStyle=modern
OutputDir=dist
OutputBaseFilename=RegKit-Setup-{#AppVersion}-{#Arch}

[Tasks]
Name: "startmenu"; Description: "Start Menu shortcut"; GroupDescription: "Shortcuts:"; Flags: checkedonce
Name: "desktopicon"; Description: "Desktop shortcut"; GroupDescription: "Shortcuts:"
Name: "replace_regedit"; Description: "Replace Regedit"; GroupDescription: "Integration:"; Check: IsAdminInstallMode

[Files]
Source: "{#BuildDir}\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildDir}\assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#BuildDir}\records\23H2.txt"; DestDir: "{app}\records"; Flags: ignoreversion
Source: "{#BuildDir}\records\24H2.txt"; DestDir: "{app}\records"; Flags: ignoreversion
Source: "{#BuildDir}\records\25H2.txt"; DestDir: "{app}\records"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\RegKit\RegKit"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"; Tasks: startmenu
Name: "{autodesktop}\RegKit"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Registry]
Root: HKLM; Subkey: "Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\regedit.exe"; ValueType: string; ValueName: "Debugger"; ValueData: """{app}\{#AppExeName}"""; Flags: uninsdeletevalue uninsdeletekeyifempty; Tasks: replace_regedit

[Code]
const
  IfeoRegeditKey = 'Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\regedit.exe';

procedure RemoveRegeditReplacement;
var
  Debugger: string;
begin
  if not RegQueryStringValue(HKLM, IfeoRegeditKey, 'Debugger', Debugger) then
    exit;
  if CompareText(RemoveQuotes(Trim(Debugger)), ExpandConstant('{app}\{#AppExeName}')) = 0 then begin
    RegDeleteValue(HKLM, IfeoRegeditKey, 'Debugger');
    RegDeleteKeyIfEmpty(HKLM, IfeoRegeditKey);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
    RemoveRegeditReplacement;
end;

function GetDefaultDir(Param: string): string;
begin
  if IsAdminInstallMode then begin
    Result := ExpandConstant('{autopf}\Noverse\RegKit');
  end else begin
    Result := ExpandConstant('{localappdata}\Noverse\RegKit');
  end;
end;

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch RegKit"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{localappdata}\Noverse\RegKit"
