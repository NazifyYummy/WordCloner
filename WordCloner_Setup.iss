#define AppVersion "1.5.1"
#define AppName "WordCloner"
#define AppExe "wordcloner.exe"
#define AppSupport "M.J.Nazify.Yummy@gmail.com"
#define AppURL "https://github.com/NazifyYummy/WordCloner"

[Setup]
AppId={{62CD052B-F3AC-48EA-AEF5-DB9AC501B5CD}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher=WordCloner Project
AppPublisherURL={#AppURL}
AppSupportURL=mailto:{#AppSupport}
AppUpdatesURL={#AppURL}

DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible

ChangesEnvironment=yes

Uninstallable=yes
CreateUninstallRegKey=yes
UninstallDisplayName={#AppName}
UninstallDisplayVersion={#AppVersion}
UninstallDisplayIcon={app}\{#AppExe}

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
Source: "dist\WordCloner\*"; DestDir: "{app}"; \
    Flags: ignoreversion recursesubdirs createallsubdirs

Source: "vc_redist.x64.exe"; DestDir: "{tmp}"; \
    Flags: deleteafterinstall skipifsourcedoesntexist

[Icons]
Name: "{group}\{#AppName}";           Filename: "{app}\{#AppExe}"; WorkingDir: "{app}"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}";     Filename: "{app}\{#AppExe}"; WorkingDir: "{app}"; \
    IconFilename: "{app}\app.ico"; Tasks: desktopicon

[Run]
Filename: "{tmp}\vc_redist.x64.exe"; \
    Parameters: "/install /quiet /norestart"; \
    Check: VC2019RedistNeedsInstall('x64'); \
    Flags: waituntilterminated; \
    StatusMsg: "Installing Visual C++ Redistributable / نصب پیش‌نیاز Microsoft Visual C++..."

Filename: "{app}\{#AppExe}"; \
    Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; \
    Flags: nowait postinstall skipifsilent

[Code]

const
  SystemEnvironmentKey = 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment';
  SystemPathValueName = 'Path';
  AppRegistryKey = 'SOFTWARE\WordCloner';
  InstallerOwnedPathValueName = 'InstallerAddedPath';
  InstallerSetHomeMarker = 'InstallerSetWordClonerHome';
  WordClonerHomeValueName = 'WORDCLONER_HOME';

function NormalizePathEntry(const PathEntry: String): String;
var
  S: String;
begin
  S := Trim(PathEntry);
  if Length(S) >= 2 then
  begin
    if (S[1] = '"') and (S[Length(S)] = '"') then
      S := Copy(S, 2, Length(S) - 2);
  end;
  S := Trim(S);
  while Length(S) > 3 do
  begin
    if S[Length(S)] <> '\' then
    begin
      Result := S;
      Exit;
    end;
    Delete(S, Length(S), 1);
  end;
  Result := S;
end;

function SamePathEntry(const FirstPath, SecondPath: String): Boolean;
begin
  Result := CompareText(
    NormalizePathEntry(FirstPath), NormalizePathEntry(SecondPath)) = 0;
end;

function PathContainsEntry(const FullPath, EntryToFind: String): Boolean;
var
  StartAt, I: Integer;
  Entry: String;
begin
  Result := False;
  StartAt := 1;
  for I := 1 to Length(FullPath) do
  begin
    if FullPath[I] = ';' then
    begin
      Entry := Copy(FullPath, StartAt, I - StartAt);
      if SamePathEntry(Entry, EntryToFind) then
      begin
        Result := True;
        Exit;
      end;
      StartAt := I + 1;
    end;
  end;
  Entry := Copy(FullPath, StartAt, Length(FullPath) - StartAt + 1);
  Result := SamePathEntry(Entry, EntryToFind);
end;

procedure AppendPathEntry(var OutputPath: String; var IsFirstEntry: Boolean;
  const Entry: String);
begin
  if not IsFirstEntry then
    OutputPath := OutputPath + ';';
  OutputPath := OutputPath + Entry;
  IsFirstEntry := False;
end;

function RemoveFirstPathEntry(const FullPath, EntryToRemove: String): String;
var
  StartAt, I: Integer;
  Entry: String;
  Removed, IsFirstEntry: Boolean;
begin
  Result := '';
  StartAt := 1;
  Removed := False;
  IsFirstEntry := True;
  for I := 1 to Length(FullPath) do
  begin
    if FullPath[I] = ';' then
    begin
      Entry := Copy(FullPath, StartAt, I - StartAt);
      if (not Removed) and SamePathEntry(Entry, EntryToRemove) then
        Removed := True
      else
        AppendPathEntry(Result, IsFirstEntry, Entry);
      StartAt := I + 1;
    end;
  end;
  Entry := Copy(FullPath, StartAt, Length(FullPath) - StartAt + 1);
  if (not Removed) and SamePathEntry(Entry, EntryToRemove) then
    Removed := True
  else
    AppendPathEntry(Result, IsFirstEntry, Entry);
end;

function WriteSystemPath(const SystemPath: String): Boolean;
begin
  Result := RegWriteExpandStringValue(
    HKEY_LOCAL_MACHINE, SystemEnvironmentKey, SystemPathValueName, SystemPath);
end;

procedure AddAppDirectoryToSystemPath(const AppDirectory: String);
var
  CurrentPath, PreviousInstallerPath, UpdatedPath: String;
  HasPreviousInstallerPath, HasCurrentPath: Boolean;
begin
  HasPreviousInstallerPath := RegQueryStringValue(
    HKEY_LOCAL_MACHINE, AppRegistryKey, InstallerOwnedPathValueName,
    PreviousInstallerPath);

  if HasPreviousInstallerPath then
  begin
    if not SamePathEntry(PreviousInstallerPath, AppDirectory) then
    begin
      if RegQueryStringValue(
           HKEY_LOCAL_MACHINE, SystemEnvironmentKey, SystemPathValueName, CurrentPath) then
      begin
        if PathContainsEntry(CurrentPath, PreviousInstallerPath) then
        begin
          UpdatedPath := RemoveFirstPathEntry(CurrentPath, PreviousInstallerPath);
          if not WriteSystemPath(UpdatedPath) then
          begin
            Log('Could not remove the previous WordCloner entry from the system PATH.');
            Exit;
          end;
          CurrentPath := UpdatedPath;
        end;
      end
      else if RegValueExists(
           HKEY_LOCAL_MACHINE, SystemEnvironmentKey, SystemPathValueName) then
      begin
        Log('Could not read the system PATH; leaving it unchanged.');
        Exit;
      end;

      if not RegDeleteValue(
           HKEY_LOCAL_MACHINE, AppRegistryKey, InstallerOwnedPathValueName) then
        Log('Could not update the stored WordCloner PATH marker.');
    end;
  end;

  HasCurrentPath := RegQueryStringValue(
    HKEY_LOCAL_MACHINE, SystemEnvironmentKey, SystemPathValueName, CurrentPath);
  if not HasCurrentPath then
  begin
    if RegValueExists(HKEY_LOCAL_MACHINE, SystemEnvironmentKey, SystemPathValueName) then
    begin
      Log('The system PATH is not a readable string; leaving it unchanged.');
      Exit;
    end;
    CurrentPath := '';
  end;

  if not PathContainsEntry(CurrentPath, AppDirectory) then
  begin
    if CurrentPath = '' then
      UpdatedPath := AppDirectory
    else if CurrentPath[Length(CurrentPath)] = ';' then
      UpdatedPath := CurrentPath + AppDirectory
    else
      UpdatedPath := CurrentPath + ';' + AppDirectory;

    if not WriteSystemPath(UpdatedPath) then
    begin
      Log('Could not add the WordCloner installation directory to the system PATH.');
      Exit;
    end;

    if not RegWriteStringValue(
         HKEY_LOCAL_MACHINE, AppRegistryKey, InstallerOwnedPathValueName,
         AppDirectory) then
      Log('Added WordCloner to the system PATH but could not record its ownership marker.');
  end;
end;

procedure RemoveInstallerOwnedPathEntry;
var
  OwnedPath, CurrentPath, UpdatedPath: String;
begin
  if not RegQueryStringValue(
       HKEY_LOCAL_MACHINE, AppRegistryKey, InstallerOwnedPathValueName, OwnedPath) then
    Exit;

  if RegQueryStringValue(
       HKEY_LOCAL_MACHINE, SystemEnvironmentKey, SystemPathValueName, CurrentPath) then
  begin
    if PathContainsEntry(CurrentPath, OwnedPath) then
    begin
      UpdatedPath := RemoveFirstPathEntry(CurrentPath, OwnedPath);
      if not WriteSystemPath(UpdatedPath) then
      begin
        Log('Could not remove the WordCloner entry from the system PATH during uninstall.');
        Exit;
      end;
    end;
  end
  else if RegValueExists(HKEY_LOCAL_MACHINE, SystemEnvironmentKey, SystemPathValueName) then
  begin
    Log('Could not read the system PATH during uninstall; leaving it unchanged.');
    Exit;
  end;

  if not RegDeleteValue(
       HKEY_LOCAL_MACHINE, AppRegistryKey, InstallerOwnedPathValueName) then
    Log('Could not remove the stored WordCloner PATH marker.');
end;

procedure SetWordClonerHomeIfMissing;
var
  Existing: String;
  DataDir: String;
begin
  if RegQueryStringValue(
       HKEY_LOCAL_MACHINE, SystemEnvironmentKey, WordClonerHomeValueName, Existing) then
  begin
    Log('WORDCLONER_HOME already set to: ' + Existing + ' (leaving unchanged)');
    Exit;
  end;

  DataDir := ExpandConstant('{commonappdata}\{#AppName}');

  if RegWriteExpandStringValue(
       HKEY_LOCAL_MACHINE, SystemEnvironmentKey, WordClonerHomeValueName, DataDir) then
  begin
    Log('Set WORDCLONER_HOME = ' + DataDir);
    RegWriteStringValue(
      HKEY_LOCAL_MACHINE, AppRegistryKey, InstallerSetHomeMarker, '1');
  end
  else
    Log('Warning: Could not set WORDCLONER_HOME');
end;

procedure RemoveWordClonerHomeIfWeSetIt;
var
  Marker: String;
begin
  if not RegQueryStringValue(
       HKEY_LOCAL_MACHINE, AppRegistryKey, InstallerSetHomeMarker, Marker) then
    Exit;

  if RegDeleteValue(
       HKEY_LOCAL_MACHINE, SystemEnvironmentKey, WordClonerHomeValueName) then
    Log('Removed WORDCLONER_HOME (set by installer)')
  else
    Log('Could not remove WORDCLONER_HOME');

  RegDeleteValue(HKEY_LOCAL_MACHINE, AppRegistryKey, InstallerSetHomeMarker);
end;

procedure CreateDataFolder;
var
  DataDir: String;
  ResultCode: Integer;
begin
  DataDir := ExpandConstant('{commonappdata}\{#AppName}');

  if not DirExists(DataDir) then
  begin
    if ForceDirectories(DataDir) then
      Log('Created data folder: ' + DataDir)
    else
    begin
      Log('Warning: Could not create data folder: ' + DataDir);
      Exit;
    end;
  end;

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

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    AddAppDirectoryToSystemPath(ExpandConstant('{app}'));
    SetWordClonerHomeIfMissing;
    CreateDataFolder;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  DataDir: String;
begin
  if CurUninstallStep = usUninstall then
  begin
    RemoveInstallerOwnedPathEntry;

    RemoveWordClonerHomeIfWeSetIt;

    if RegValueExists(HKEY_CURRENT_USER,
         'Software\Microsoft\Windows\CurrentVersion\Run', 'WordCloner') then
    begin
      if not RegDeleteValue(HKEY_CURRENT_USER,
           'Software\Microsoft\Windows\CurrentVersion\Run', 'WordCloner') then
        Log('Could not remove the WordCloner startup registration.');
    end;

    RegDeleteKeyIfEmpty(HKEY_LOCAL_MACHINE, AppRegistryKey);
  end;

  if CurUninstallStep = usPostUninstall then
  begin
    DataDir := ExpandConstant('{commonappdata}\{#AppName}');
    if DirExists(DataDir) then
    begin
      if MsgBox(
           'Do you also want to delete all WordCloner data?' + #13#10 +
           'This will remove settings.json, templates.json, and the Words folder' + #13#10 +
           'shared by ALL users of this computer.' + #13#10#13#10 +
           'آیا داده‌های WordCloner (مشترک بین همه کاربران) هم حذف شوند؟',
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