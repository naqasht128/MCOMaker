; MCO Maker - Inno Setup Installer Script
; Produces a modern, professional, per-user Windows Setup Wizard

#ifndef AppVersion
#define AppVersion "1.0.0"
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

[Files]
Source: "dist\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "redist\api-ms-win-core-path-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "templates\*"; DestDir: "{app}\templates"; Flags: ignoreversion recursesubdirs createallsubdirs
#if FileExists("redist\vc_redist.x64.exe")
Source: "redist\vc_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall
#endif
#if FileExists("redist\vc_redist.x86.exe")
Source: "redist\vc_redist.x86.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall
#endif

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\assets\mcomaker.ico"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\assets\mcomaker.ico"; Tasks: desktopicon

[Run]
#if FileExists("redist\vc_redist.x64.exe")
Filename: "{tmp}\vc_redist.x64.exe"; Parameters: "/install /passive /norestart"; Check: VCRedistNeedsInstallX64; Flags: waituntilterminated; StatusMsg: "Configuring Microsoft Visual C++ Runtime (x64)..."
#endif
#if FileExists("redist\vc_redist.x86.exe")
Filename: "{tmp}\vc_redist.x86.exe"; Parameters: "/install /passive /norestart"; Check: VCRedistNeedsInstallX86; Flags: waituntilterminated; StatusMsg: "Configuring Microsoft Visual C++ Runtime (x86)..."
#endif
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent shellexec; WorkingDir: "{app}"

[Code]
function SetEnvironmentVariable(lpName, lpValue: string): BOOL;
external 'SetEnvironmentVariableW@kernel32.dll stdcall';

function VCRedistNeedsInstallX64: Boolean;
var
  Installed: Cardinal;
begin
  Result := True;
  if RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\X64', 'Installed', Installed) and (Installed = 1) then
    Result := False
  else if RegQueryDWordValue(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\X64', 'Installed', Installed) and (Installed = 1) then
    Result := False;
end;

function VCRedistNeedsInstallX86: Boolean;
var
  Installed: Cardinal;
begin
  Result := True;
  if RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\X86', 'Installed', Installed) and (Installed = 1) then
    Result := False
  else if RegQueryDWordValue(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\X86', 'Installed', Installed) and (Installed = 1) then
    Result := False;
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
