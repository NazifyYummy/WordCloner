#define AppVersion "2.1.0"
#define MyAppName "WordCloner"
#define MyAppExeName "wordcloner.exe"
#define MyAppPublisher "WordCloner Project"
#define MyAppURL "https://github.com/NazifyYummy/WordCloner"
#define MyAppSupport "M.J.Nazify.Yummy@gmail.com"

[Setup]
AppId={{62CD052B-F3AC-48EA-AEF5-DB9AC501B5CD}
AppName={#MyAppName}
AppVersion={#AppVersion}
AppVerName={#MyAppName} {#AppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppSupport}
AppUpdatesURL={#MyAppURL}

DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible

ChangesEnvironment=yes

Uninstallable=yes
CreateUninstallRegKey=yes
UninstallDisplayName={#MyAppName} {#AppVersion}
UninstallDisplayIcon={app}\{#MyAppExeName}

OutputDir=installer
OutputBaseFilename=WordCloner_{#AppVersion}_Setup_64bit
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes

SetupIconFile=src\app.ico

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; \
    GroupDescription: "{cm:AdditionalIcons}"; Flags: checkedonce

[Files]
Source: "src\dist\WordCloner\*"; DestDir: "{app}"; \
    Flags: ignoreversion recursesubdirs createallsubdirs

Source: "vc_redist.x64.exe"; DestDir: "{tmp}"; \
    Flags: deleteafterinstall skipifsourcedoesntexist

[Icons]
Name: "{group}\{#MyAppName}";           Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}";     Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; \
    IconFilename: "{app}\app.ico"; Tasks: desktopicon

[Registry]
Root: HKLM; \
    Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; \
    ValueType: expandsz; ValueName: "Path"; \
    ValueData: "{olddata};{app}"; \
    Check: NeedsAddPath('{app}')

Root: HKLM; \
    Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; \
    ValueType: expandsz; ValueName: "WORDCLONER_HOME"; \
    ValueData: "{commonappdata}\{#MyAppName}"; \
    Flags: uninsdeletevalue

[Run]
Filename: "{tmp}\vc_redist.x64.exe"; \
    Parameters: "/install /quiet /norestart"; \
    Check: VC2019RedistNeedsInstall('x64'); \
    Flags: waituntilterminated; \
    StatusMsg: "Installing Visual C++ Redistributable / نصب پیش‌نیاز Microsoft Visual C++..."

Filename: "{app}\{#MyAppExeName}"; \
    Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; \
    Flags: nowait postinstall skipifsilent

[Code]

function NeedsAddPath(Param: string): Boolean;
var
  OrigPath: String;
begin
  if not RegQueryStringValue(HKEY_LOCAL_MACHINE,
    'SYSTEM\CurrentControlSet\Control\Session Manager\Environment',
    'Path', OrigPath) then
  begin
    Result := True;
    exit;
  end;
  Result := Pos(';' + Param + ';', ';' + OrigPath + ';') = 0;
end;

procedure RemoveFromPath(const Dir: String);
var
  OrigPath, NewPath: String;
begin
  if not RegQueryStringValue(HKEY_LOCAL_MACHINE,
    'SYSTEM\CurrentControlSet\Control\Session Manager\Environment',
    'Path', OrigPath) then
    exit;

  NewPath := OrigPath;
  StringChangeEx(NewPath, ';' + Dir, '', True);
  StringChangeEx(NewPath, Dir + ';', '', True);
  StringChangeEx(NewPath, Dir, '', True);

  RegWriteExpandStringValue(HKEY_LOCAL_MACHINE,
    'SYSTEM\CurrentControlSet\Control\Session Manager\Environment',
    'Path', NewPath);
end;

function VC2019RedistNeedsInstall(Arch: String): Boolean;
var
  Version: String;
begin
  if RegQueryStringValue(
       HKEY_LOCAL_MACHINE,
       'SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\' + Arch,
       'Version', Version) then
  begin
    Log('VC Redist Version check : found ' + Version);
    Result := (CompareStr(Version, 'v14.26.28720.03') < 0);
  end
  else
  begin
    Result := True;
  end;
end;

procedure CreateDataFolder();
var
  DataDir: String;
  ResultCode: Integer;
begin
  DataDir := ExpandConstant('{commonappdata}\{#MyAppName}');

  if not DirExists(DataDir) then
  begin
    if ForceDirectories(DataDir) then
      Log('Created data folder: ' + DataDir)
    else
      Log('Warning: Could not create data folder: ' + DataDir);
  end;

  if DirExists(DataDir) then
  begin
    if Exec(
         ExpandConstant('{sys}\icacls.exe'),
         '"' + DataDir + '" /grant *S-1-5-32-545:(OI)(CI)M /T /C /Q',
         '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    begin
      if ResultCode = 0 then
        Log('Data folder permissions set successfully.')
      else
        Log('Warning: icacls returned code ' + IntToStr(ResultCode));
    end
    else
      Log('Warning: Could not run icacls.exe');
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    CreateDataFolder();
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  DataDir: String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    RemoveFromPath(ExpandConstant('{app}'));

    if RegValueExists(HKEY_CURRENT_USER,
         'Software\Microsoft\Windows\CurrentVersion\Run',
         'WordCloner') then
    begin
      if not RegDeleteValue(HKEY_CURRENT_USER,
           'Software\Microsoft\Windows\CurrentVersion\Run',
           'WordCloner') then
        Log('Could not remove the WordCloner startup registration.');
    end;
  end;

  if CurUninstallStep = usPostUninstall then
  begin
    DataDir := ExpandConstant('{commonappdata}\{#MyAppName}');
    if DirExists(DataDir) then
    begin
      if MsgBox(
           'Do you also want to delete all WordCloner data?' + #13#10 +
           'This will remove settings.json, templates.json, and the Words folder.' + #13#10#13#10 +
           'آیا داده‌های WordCloner هم حذف شوند؟' + #13#10 +
           '(settings.json، templates.json و پوشه Words)',
           mbConfirmation, MB_YESNO) = IDYES then
      begin
        DelTree(DataDir, True, True, True);
        Log('Data folder deleted: ' + DataDir);
      end
      else
        Log('User chose to keep data folder: ' + DataDir);
    end;
  end;
end;