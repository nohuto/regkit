#define AppId "4678f42c-c6a2-4df9-bc2a-dddbd2613045"
#define AppName "RegKit"
#define AppExeName "regkit.exe"
#define AppPublisher "nohuto"
#define AppCopyright "(C) 2026 nohuto"
#define AppURL "https://github.com/nohuto/regkit"
#ifndef Arch
  #define Arch "x64"
#endif
#if Arch == "x86"
  #define BuildDir "..\\build\\x86-release\\Release"
#else
  #define BuildDir "..\\build\\x64-release\\Release"
#endif
#define AppVersion GetFileVersionString(AddBackslash(BuildDir) + AppExeName)

[Setup]
AppId={#AppId}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
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
WizardSmallImageFile=images\small-55.png,images\small-69.png,images\small-83.png,images\small-97.png,images\small-110.png,images\small-138.png
WizardImageFile=images\large-100.png,images\large-125.png,images\large-150.png,images\large-175.png,images\large-200.png,images\large-250.png
WizardSmallImageFileDynamicDark=images\small-55.png,images\small-69.png,images\small-83.png,images\small-97.png,images\small-110.png,images\small-138.png
WizardImageFileDynamicDark=images\large-100.png,images\large-125.png,images\large-150.png,images\large-175.png,images\large-200.png,images\large-250.png
Compression=lzma2
SolidCompression=yes
ChangesAssociations=yes
#if Arch == "x64"
ArchitecturesAllowed=x64os
ArchitecturesInstallIn64BitMode=x64os
#endif
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog
WizardStyle=modern dynamic
CloseApplications=yes
RestartApplications=no
OutputDir=..\build\dist
OutputBaseFilename=RegKit-Setup-{#AppVersion}-{#Arch}

[Messages]
PrivilegesRequiredOverrideTitle=Install Mode
PrivilegesRequiredOverrideInstruction=Select install mode
PrivilegesRequiredOverrideText1=Install %1 for:
PrivilegesRequiredOverrideAllUsersRecommended=&All users
PrivilegesRequiredOverrideCurrentUser=&Only me

[Types]
Name: "full"; Description: "Full installation"
Name: "minimal"; Description: "Minimal installation"
Name: "custom"; Description: "Custom installation"; Flags: iscustom

[Components]
Name: "main"; Description: "RegKit"; Types: full minimal custom; Flags: fixed
Name: "icons"; Description: "Classic toolbar icon set"; Types: full
Name: "comments"; Description: "Default comments"; Types: full
Name: "traces"; Description: "Trace files for the Trace menu"; Types: full
Name: "bitfields"; Description: "Bitfield definitions"; Types: full
Name: "defaults"; Description: "Registry exports for the Default menu"; Types: full
Name: "lang"; Description: "Languages"
Name: "lang\de"; Description: "Deutsch"; Types: full; Check: SystemLanguage(7)
Name: "lang\de_other"; Description: "Deutsch"; Check: not SystemLanguage(7)
Name: "lang\fr"; Description: "Français"; Types: full; Check: SystemLanguage(12)
Name: "lang\fr_other"; Description: "Français"; Check: not SystemLanguage(12)
Name: "lang\ru"; Description: "Русский"; Types: full; Check: SystemLanguage(25)
Name: "lang\ru_other"; Description: "Русский"; Check: not SystemLanguage(25)
Name: "lang\es"; Description: "Español"; Types: full; Check: SystemLanguage(10)
Name: "lang\es_other"; Description: "Español"; Check: not SystemLanguage(10)
Name: "lang\it"; Description: "Italiano"; Types: full; Check: SystemLanguage(16)
Name: "lang\it_other"; Description: "Italiano"; Check: not SystemLanguage(16)
Name: "lang\ja"; Description: "日本語"; Types: full; Check: SystemLanguage(17)
Name: "lang\ja_other"; Description: "日本語"; Check: not SystemLanguage(17)
Name: "lang\pt_BR"; Description: "Português (Brasil)"; Types: full; Check: SystemLanguage(22)
Name: "lang\pt_BR_other"; Description: "Português (Brasil)"; Check: not SystemLanguage(22)
Name: "lang\zh_CN"; Description: "简体中文"; Types: full; Check: SystemLanguage(4)
Name: "lang\zh_CN_other"; Description: "简体中文"; Check: not SystemLanguage(4)
Name: "lang\ko"; Description: "한국어"; Types: full; Check: SystemLanguage(18)
Name: "lang\ko_other"; Description: "한국어"; Check: not SystemLanguage(18)

[Tasks]
Name: "startmenu"; Description: "Start Menu shortcut"; GroupDescription: "Shortcuts:"; Flags: checkedonce
Name: "desktopicon"; Description: "Desktop shortcut"; GroupDescription: "Shortcuts:"
Name: "replace_regedit"; Description: "Replace RegEdit"; GroupDescription: "Integration:"; Check: IsAdminInstallMode
Name: "edit_context_menu"; Description: "Add ""Edit"" Context Menu"; GroupDescription: "Integration:"; Flags: checkedonce
Name: "auto_updates"; Description: "Check for updates automatically"; GroupDescription: "Options:"; Check: not IsUpgrade
Name: "simple_mode"; Description: "Enable simple mode (RegEdit layout)"; GroupDescription: "Options:"; Flags: unchecked; Check: not IsUpgrade

[Files]
Source: "{#BuildDir}\offreg.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "..\redist\pcre2\LICENCE.md"; DestDir: "{app}\licences"; DestName: "PCRE2-LICENCE.md"; Flags: ignoreversion; Components: main
Source: "{#BuildDir}\assets\icons\classic\*"; DestDir: "{app}\assets\icons\classic"; Flags: ignoreversion; Components: icons
Source: "{#BuildDir}\assets\comments\*"; DestDir: "{app}\assets\comments"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: comments
Source: "{#BuildDir}\assets\records\*"; DestDir: "{app}\assets\records"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: traces
Source: "{#BuildDir}\assets\bitfields\*"; DestDir: "{app}\assets\bitfields"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: bitfields
Source: "{#BuildDir}\assets\lang\de.po"; DestDir: "{app}\assets\lang"; Flags: ignoreversion; Components: lang\de lang\de_other
Source: "{#BuildDir}\assets\lang\fr.po"; DestDir: "{app}\assets\lang"; Flags: ignoreversion; Components: lang\fr lang\fr_other
Source: "{#BuildDir}\assets\lang\ru.po"; DestDir: "{app}\assets\lang"; Flags: ignoreversion; Components: lang\ru lang\ru_other
Source: "{#BuildDir}\assets\lang\es.po"; DestDir: "{app}\assets\lang"; Flags: ignoreversion; Components: lang\es lang\es_other
Source: "{#BuildDir}\assets\lang\it.po"; DestDir: "{app}\assets\lang"; Flags: ignoreversion; Components: lang\it lang\it_other
Source: "{#BuildDir}\assets\lang\ja.po"; DestDir: "{app}\assets\lang"; Flags: ignoreversion; Components: lang\ja lang\ja_other
Source: "{#BuildDir}\assets\lang\pt-BR.po"; DestDir: "{app}\assets\lang"; Flags: ignoreversion; Components: lang\pt_BR lang\pt_BR_other
Source: "{#BuildDir}\assets\lang\zh-CN.po"; DestDir: "{app}\assets\lang"; Flags: ignoreversion; Components: lang\zh_CN lang\zh_CN_other
Source: "{#BuildDir}\assets\lang\ko.po"; DestDir: "{app}\assets\lang"; Flags: ignoreversion; Components: lang\ko lang\ko_other
Source: "{#BuildDir}\assets\defaults\*"; DestDir: "{app}\assets\defaults"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: defaults
Source: "{#BuildDir}\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion; Components: main

[InstallDelete]
Type: files; Name: "{app}\is-*.tmp"

[Icons]
Name: "{autoprograms}\RegKit\RegKit"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"; Tasks: startmenu
Name: "{autodesktop}\RegKit"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Registry]
Root: HKA; Subkey: "Software\Classes\Applications\{#AppExeName}"; ValueType: string; ValueName: "FriendlyAppName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\Applications\{#AppExeName}\DefaultIcon"; ValueType: string; ValueData: """{app}\{#AppExeName}"",0"
Root: HKA; Subkey: "Software\Classes\Applications\{#AppExeName}\SupportedTypes"; ValueType: string; ValueName: ".reg"; ValueData: ""
Root: HKA; Subkey: "Software\Classes\Applications\{#AppExeName}\shell\open\command"; ValueType: string; ValueData: """{app}\{#AppExeName}"" ""%1"""
[Code]
// full installs only the pack of the Windows display language
function SystemLanguage(PrimaryLanguage: Integer): Boolean;
begin
  Result := (GetUILanguage and $3FF) = PrimaryLanguage;
end;

function ConfirmOverride: Boolean;
var
  Form: TSetupForm;
  Text: TNewStaticText;
  OverrideButton, CancelButton: TNewButton;
  W: Integer;
begin
  Result := False;
  if WizardSilent then
    exit;
  Form := CreateCustomForm(ScaleX(380), ScaleY(100), False, False);
  try
    Form.Caption := 'Replace RegEdit';
    Text := TNewStaticText.Create(Form);
    Text.Parent := Form;
    Text.AutoSize := False;
    Text.WordWrap := True;
    Text.Left := ScaleX(12);
    Text.Top := ScaleY(12);
    Text.Width := Form.ClientWidth - ScaleX(24);
    Text.Height := Form.ClientHeight - ScaleY(12 + 23 + 24);
    Text.Caption := 'RegEdit already has a Debugger entry owned by another program.'#13#10#13#10'Override the existing entry?';

    OverrideButton := TNewButton.Create(Form);
    OverrideButton.Parent := Form;
    OverrideButton.Caption := 'Override';
    OverrideButton.ModalResult := mrOk;

    CancelButton := TNewButton.Create(Form);
    CancelButton.Parent := Form;
    CancelButton.Caption := 'Cancel';
    CancelButton.ModalResult := mrCancel;
    CancelButton.Cancel := True;
    CancelButton.Default := True;

    W := Form.CalculateButtonWidth([OverrideButton.Caption, CancelButton.Caption]);
    CancelButton.SetBounds(Form.ClientWidth - ScaleX(12) - W, Form.ClientHeight - ScaleY(12 + 23), W, ScaleY(23));
    OverrideButton.SetBounds(CancelButton.Left - ScaleX(6) - W, CancelButton.Top, W, ScaleY(23));

    Form.ActiveControl := CancelButton;
    Form.FlipAndCenterIfNeeded(True, WizardForm, False);
    Result := Form.ShowModal = mrOk;
  finally
    Form.Free;
  end;
end;

procedure InstallRegEditReplacement;
var
  ResultCode: Integer;
begin
  if not WizardIsTaskSelected('replace_regedit') then
    exit;
  if not Exec(ExpandConstant('{app}\{#AppExeName}'), '--install-regedit-replacement', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    ResultCode := 1;
  if ResultCode = 2 then begin
    if not ConfirmOverride then
      exit;
    if not Exec(ExpandConstant('{app}\{#AppExeName}'), '--install-regedit-replacement --override', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
      ResultCode := 1;
  end;
  if ResultCode <> 0 then
    SuppressibleMsgBox('RegEdit replacement couldn''t be installed. Non administrators may be able to modify the install folder.', mbError, MB_OK, IDOK);
end;

procedure RemoveRegEditReplacement;
var
  ResultCode: Integer;
begin
  if not Exec(ExpandConstant('{app}\{#AppExeName}'), '--uninstall-regedit-replacement', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) or (ResultCode <> 0) then begin
    RaiseException('RegEdit replacement couldn''t be removed. Uninstallation was stopped to avoid leaving RegEdit redirected to a deleted file.');
  end;
end;

const
  UninstallKey = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#AppId}_is1';

var
  Upgrade, Installing: Boolean;
  DataAside: String;

procedure RemoveEmptyUninstallLogs;
var
  Folder: String;
  Find: TFindRec;
begin
  Folder := AddBackslash(ExpandConstant('{app}'));
  if not FindFirst(Folder + 'unins???.dat', Find) then
    exit;
  try
    repeat
      if (Find.SizeHigh = 0) and (Find.SizeLow = 0) then begin
        Log('removing empty uninstall log ' + Find.Name);
        DeleteFile(Folder + Find.Name);
        DeleteFile(Folder + ChangeFileExt(Find.Name, '.exe'));
        DeleteFile(Folder + ChangeFileExt(Find.Name, '.msg'));
      end;
    until not FindNext(Find);
  finally
    FindClose(Find);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then begin
    Installing := True;
    RemoveEmptyUninstallLogs;
  end;
  if CurStep = ssPostInstall then begin
    InstallRegEditReplacement;
  end;
end;

procedure DeinitializeSetup;
var
  ResultCode: Integer;
begin
  if Installing and WizardSilent and (ExpandConstant('{param:RELAUNCH|0}') = '1') then
    ExecAsOriginalUser(ExpandConstant('{app}\{#AppExeName}'), '--restore-session', '', SW_SHOWNORMAL, ewNoWait, ResultCode);
end;

function DataFolder: String;
begin
  Result := ExpandConstant('{localappdata}\Noverse\RegKit');
end;

function IsProgramEntry(const Name: String): Boolean;
begin
  Result := False;
  if CompareText(RemoveBackslash(ExpandConstant('{app}')), DataFolder) <> 0 then
    exit;
  Result := (CompareText(Name, '{#AppExeName}') = 0) or (CompareText(Name, 'offreg.dll') = 0) or (CompareText(Name, 'licences') = 0) or
    (CompareText(Name, 'assets') = 0) or WildcardMatch(Lowercase(Name), 'unins*') or WildcardMatch(Lowercase(Name), 'is-*.tmp');
end;

function ListEntries(const Folder: String; SkipProgram: Boolean): TStringList;
var
  Find: TFindRec;
begin
  Result := TStringList.Create;
  if not FindFirst(AddBackslash(Folder) + '*', Find) then
    exit;
  try
    repeat
      if (Find.Name <> '.') and (Find.Name <> '..') then
        if not (SkipProgram and IsProgramEntry(Find.Name)) then
          Result.Add(Find.Name);
    until not FindNext(Find);
  finally
    FindClose(Find);
  end;
end;

procedure MoveEntries(const Source, Dest: String; SkipProgram: Boolean);
var
  Names: TStringList;
  Target: String;
  I: Integer;
begin
  Names := ListEntries(Source, SkipProgram);
  try
    for I := 0 to Names.Count - 1 do begin
      Target := AddBackslash(Dest) + Names.Strings[I];
      if DirExists(Target) then
        DelTree(Target, True, True, True)
      else if FileExists(Target) then
        DeleteFile(Target);
      if not RenameFile(AddBackslash(Source) + Names.Strings[I], Target) then
        Log('could not move ' + Names.Strings[I] + ' to ' + Dest);
    end;
  finally
    Names.Free;
  end;
end;

procedure SetDataAside;
var
  Names: TStringList;
  Count: Integer;
begin
  Names := ListEntries(DataFolder, True);
  Count := Names.Count;
  Names.Free;
  if Count = 0 then
    exit;
  if not CreateDir(DataFolder + '.uninstall') then
    exit;
  MoveEntries(DataFolder, DataFolder + '.uninstall', True);
  DataAside := DataFolder + '.uninstall';
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Remove: Boolean;
begin
  if CurUninstallStep = usUninstall then begin
    RemoveRegEditReplacement;
    SetDataAside;
  end;
  if CurUninstallStep <> usPostUninstall then
    exit;
  if DataAside <> '' then begin
    Remove := False;
    if not UninstallSilent then
      Remove := MsgBox('Also remove RegKit''s data? This deletes your settings, favorites, theme presets, custom icons and history.', mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES;
    if Remove then
      DelTree(DataAside, True, True, True)
    else begin
      ForceDirectories(DataFolder);
      MoveEntries(DataAside, DataFolder, False);
      RemoveDir(DataAside);
    end;
  end;
  RemoveDir(DataFolder);
  RemoveDir(ExpandConstant('{localappdata}\Noverse'));
end;

function InitializeSetup: Boolean;
var
  Installed: String;
  InstalledVersion, ThisVersion: Int64;
begin
  Result := True;
  Upgrade := RegValueExists(HKA, UninstallKey, 'UninstallString');
  if not RegQueryStringValue(HKA, UninstallKey, 'DisplayVersion', Installed) then
    exit;
  if StrToVersion(Installed, InstalledVersion) and StrToVersion('{#AppVersion}', ThisVersion) then
    if ComparePackedVersion(InstalledVersion, ThisVersion) <= 0 then
      exit;
  Log('refused: installed version ' + Installed + ' is newer than {#AppVersion} or unreadable');
  if not WizardSilent then
    MsgBox('RegKit ' + Installed + ' is already installed. This setup installs the older version {#AppVersion} and stops.', mbError, MB_OK);
  Result := False;
end;

function IsUpgrade: Boolean;
begin
  Result := Upgrade;
end;

function TaskValue(Task: string): string;
begin
  if WizardIsTaskSelected(Task) then
    Result := '1'
  else
    Result := '0';
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
Filename: "{app}\{#AppExeName}"; Parameters: "--install-edit-context-menu"; Flags: runhidden runasoriginaluser; Tasks: edit_context_menu
Filename: "{app}\{#AppExeName}"; Parameters: "--setting auto_check_updates={code:TaskValue|auto_updates} --setting view_simple_mode={code:TaskValue|simple_mode}"; Flags: runhidden runasoriginaluser; Check: not IsUpgrade
Filename: "{app}\{#AppExeName}"; Description: "Launch RegKit"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{app}\{#AppExeName}"; Parameters: "--uninstall-edit-context-menu"; RunOnceId: "RemoveEditContextMenu"; Flags: runhidden

[UninstallDelete]
Type: files; Name: "{app}\is-*.tmp"
Type: dirifempty; Name: "{app}"
Type: dirifempty; Name: "{autopf}\Noverse"
Type: dirifempty; Name: "{localappdata}\Noverse"
