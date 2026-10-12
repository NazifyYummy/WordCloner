; ============================================================
;  WordCloner Installer Script — Version 2.0.0
;  Built with Inno Setup 6+
; ============================================================

#define MyAppName "WordCloner"
#define MyAppVersion "2.0.0"
#define MyAppPublisher "Mohammad Javad Nazify Yummy"
#define MyAppURL "https://github.com/NazifyYummy/WordCloner"
#define MyAppExeName "wordcloner.exe"
#define MyAppSupport "M.J.Nazify.Yummy@gmail.com"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppSupport}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=WordCloner-2.0.0-setup-64bit
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ChangesEnvironment=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
SetupIconFile=src\app.ico

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; \
    GroupDescription: "{cm:AdditionalIcons}"; Flags: checkedonce

[Files]
Source: "src\dist\WordCloner\*"; DestDir: "{app}"; \
    Flags: ignoreversion recursesubdirs createallsubdirs
Source: "vc_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; \
    Tasks: desktopicon; IconFilename: "{app}\app.ico"

[Registry]
Root: HKLM; \
    Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; \
    ValueType: expandsz; ValueName: "Path"; \
    ValueData: "{olddata};{app}"; \
    Check: NeedsAddPath('{app}')

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

function NeedsAddPath(Param: string): boolean;
var
  OrigPath: string;
begin
  if not RegQueryStringValue(HKEY_LOCAL_MACHINE,
    'SYSTEM\CurrentControlSet\Control\Session Manager\Environment',
    'Path', OrigPath)
  then begin
    Result := True;
    exit;
  end;
  Result := Pos(';' + Param + ';', ';' + OrigPath + ';') = 0;
end;

procedure RemoveFromPath(const Dir: string);
var
  OrigPath, NewPath: string;
begin
  if not RegQueryStringValue(HKEY_LOCAL_MACHINE,
    'SYSTEM\CurrentControlSet\Control\Session Manager\Environment',
    'Path', OrigPath)
  then exit;

  NewPath := OrigPath;
  StringChangeEx(NewPath, ';' + Dir, '', True);
  StringChangeEx(NewPath, Dir + ';', '', True);
  StringChangeEx(NewPath, Dir, '', True);

  RegWriteExpandStringValue(HKEY_LOCAL_MACHINE,
    'SYSTEM\CurrentControlSet\Control\Session Manager\Environment',
    'Path', NewPath);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
    RemoveFromPath(ExpandConstant('{app}'));
end;

function VC2019RedistNeedsInstall(Arch: string): Boolean;
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