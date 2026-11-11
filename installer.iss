; VM Migration Suite - Inno Setup 6
#ifndef MyAppVersion
  #define MyAppVersion "3.0.0"
#endif
#ifndef SourceDir
  #define SourceDir "dist"
#endif
#define MyAppName "VM Migration Suite"
#define MyAppPublisher "vm technology"
#define MyAppExeName "VMMigrationSuite.exe"

[Setup]
AppId={{7B1F7C2E-5A5D-4C3B-9D1E-3A0C5E7F9B21}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\VM Migration Suite
DefaultGroupName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}
OutputBaseFilename=VMMigrationSuite_Setup_{#MyAppVersion}
SetupIconFile={#SourceDir}\app.ico
Compression=lzma2/ultra64
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
MinVersion=10.0
DisableProgramGroupPage=yes

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "{#SourceDir}\{#MyAppExeName}";        DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\VMMigration.Engine.psm1"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\VMMigration.CLI.ps1";     DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\GUI.xaml";                DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\logo.png";                DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\app.ico";                 DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\README.md";               DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\Avvia-CLI.cmd";           DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\app.ico"
Name: "{group}\{#MyAppName} - riga di comando"; Filename: "powershell.exe"; Parameters: "-NoExit -NoProfile -ExecutionPolicy Bypass -Command ""Set-Location '{app}'; Get-Help .\VMMigration.CLI.ps1 -Examples"""; IconFilename: "{app}\app.ico"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\app.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: files; Name: "{app}\vmig_error.log"
