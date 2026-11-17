; MCO Maker - Inno Setup Installer Script
; Produces a modern, professional, per-user Windows Setup Wizard

#ifndef AppVersion
#define AppVersion "1.0.10"
#endif

#ifndef OutputBaseFilename
#define OutputBaseFilename "MCOMaker-Setup"
#endif

#define AppName "MCO Maker"
#define AppPublisher "TN Developers Co."
#define AppExeName "MCOMaker.exe"

[Setup]
AppId={{B4775642-FC32-4827-85CC-0CAA663924A2}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=mailto:tahir.naqash4@gmail.com
AppSupportURL=https://wa.me/923339962474
DefaultDirName={localappdata}\Programs\MCOMaker
DisableDirPage=yes
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename={#OutputBaseFilename}
SetupIconFile=assets\mcomaker.ico
UninstallDisplayIcon={app}\{#AppExeName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
MinVersion=10.0
PrivilegesRequired=lowest
CloseApplications=force
RestartApplications=no
VersionInfoCompany={#AppPublisher}
VersionInfoDescription={#AppName} Setup
VersionInfoVersion={#AppVersion}
VersionInfoProductName={#AppName}
VersionInfoProductVersion={#AppVersion}
VersionInfoCopyright=Copyright (C) 2026 {#AppPublisher}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[InstallDelete]
; Clean up previous monolithic onefile executable when upgrading
Type: files; Name: "{app}\{#AppExeName}"

[Files]
; The modern onedir distribution built by PyInstaller
Source: "dist\MCOMaker\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "templates\*"; DestDir: "{app}\templates"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\assets\mcomaker.ico"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\assets\mcomaker.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent shellexec; WorkingDir: "{app}"
Filename: "{app}\{#AppExeName}"; Flags: nowait shellexec; Check: ShouldLaunchSilent; WorkingDir: "{app}"

[Code]
function SetEnvironmentVariable(lpName, lpValue: string): BOOL;
external 'SetEnvironmentVariableW@kernel32.dll stdcall';

function ShouldLaunchSilent: Boolean;
begin
  Result := WizardSilent or (ExpandConstant('{param:AUTOUPDATE|0}') = '1');
end;

function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  // Clear any inherited PyInstaller variables so the relaunched app unpacks cleanly
  SetEnvironmentVariable('_MEIPASS2', '');
  SetEnvironmentVariable('_MEIPASS', '');
  SetEnvironmentVariable('PYTHONHOME', '');
  SetEnvironmentVariable('PYTHONPATH', '');

  // Terminate any running instances of MCOMaker.exe
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM MCOMaker.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(300);
  Result := True;
end;
