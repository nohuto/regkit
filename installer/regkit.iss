#define AppId "4678f42c-c6a2-4df9-bc2a-dddbd2613045"
#define AppName "RegKit"
#define AppExeName "regkit.exe"
#define AppVersion "0.0.0.1"
#define AppPublisher "Noverse (Nohuto)"
#define AppURL "https://github.com/nohuto/regkit"
#define BuildDir "..\\build\\Release"

[Setup]
AppId={#AppId}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
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
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog
WizardStyle=modern
OutputDir=dist
OutputBaseFilename=RegKit-Setup-{#AppVersion}

[Tasks]
Name: "startmenu"; Description: "Start Menu shortcut"; GroupDescription: "Shortcuts:"; Flags: checkedonce
Name: "desktopicon"; Description: "Desktop shortcut"; GroupDescription: "Shortcuts:"
Name: "contextmenu_reg"; Description: "Add 'Edit with RegKit' to .reg context menu"; GroupDescription: "File integration:"
Name: "replace_regedit"; Description: "Replace Regedit"; GroupDescription: "File integration:"

[Files]
Source: "{#BuildDir}\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildDir}\assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\RegKit\RegKit"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"; Tasks: startmenu
Name: "{userdesktop}\RegKit"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Registry]
; "Edit with RegKit" context menu for .reg
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.reg\shell\EditWithRegKit"; ValueType: string; ValueName: ""; ValueData: "Edit with RegKit"; Flags: uninsdeletekey; Tasks: contextmenu_reg
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.reg\shell\EditWithRegKit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1"""; Flags: uninsdeletekey; Tasks: contextmenu_reg

; Replace Regedit
Root: HKLM; Subkey: "Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\regedit.exe"; ValueType: string; ValueName: "Debugger"; ValueData: """{app}\{#AppExeName}"""; Flags: uninsdeletevalue; Tasks: replace_regedit

[Code]
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
