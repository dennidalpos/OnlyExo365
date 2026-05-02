#ifndef ProductName
  #define ProductName "OnlyExo365"
#endif
#ifndef Manufacturer
  #define Manufacturer "OnlyExo365"
#endif
#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#ifndef FileVersion
  #define FileVersion "0.0.0.0"
#endif
#ifndef PublishDirX64
  #error "PublishDirX64 define is required."
#endif
#ifndef OutputDir
  #error "OutputDir define is required."
#endif
#ifndef OutputBaseFilename
  #define OutputBaseFilename "OnlyExo365.Setup"
#endif
#ifndef IconPath
  #error "IconPath define is required."
#endif
#ifndef LicensePath
  #error "LicensePath define is required."
#endif

[Setup]
AppId={{B9E5A61C-8D6A-4B10-8F50-2BB72D7497F3}
AppName={#ProductName}
AppVersion={#AppVersion}
AppVerName={#ProductName} {#AppVersion}
AppPublisher={#Manufacturer}
DefaultDirName={autopf}\{#ProductName}
DefaultGroupName={#ProductName}
DisableProgramGroupPage=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
OutputDir={#OutputDir}
OutputBaseFilename={#OutputBaseFilename}
SetupIconFile={#IconPath}
LicenseFile={#LicensePath}
UninstallDisplayIcon={app}\OnlyExo365.Shell.exe
Compression=lzma2/ultra64
SolidCompression=yes
VersionInfoVersion={#FileVersion}
VersionInfoCompany={#Manufacturer}
VersionInfoDescription={#ProductName} Setup
VersionInfoProductName={#ProductName}
VersionInfoProductVersion={#AppVersion}
VersionInfoTextVersion={#FileVersion}
CloseApplications=no
ChangesEnvironment=no
AlwaysShowDirOnReadyPage=yes
ArchitecturesAllowed=x64compatible

[Files]
Source: "{#PublishDirX64}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs notimestamp; Check: Is64BitInstallMode

[Icons]
Name: "{commonprograms}\{#ProductName}"; Filename: "{app}\OnlyExo365.Shell.exe"; WorkingDir: "{app}"
Name: "{commondesktop}\{#ProductName}"; Filename: "{app}\OnlyExo365.Shell.exe"; WorkingDir: "{app}"

[Registry]
Root: HKLM; Subkey: "Software\{#ProductName}"; ValueType: string; ValueName: "InstallLocation"; ValueData: "{app}"; Flags: uninsdeletekeyifempty
Root: HKCU; Subkey: "Software\{#ProductName}"; ValueType: string; ValueName: "LogDirectory"; ValueData: "{localappdata}\{#ProductName}\logs"; Flags: uninsdeletekeyifempty
Root: HKCU; Subkey: "Software\{#ProductName}"; ValueType: string; ValueName: "SecretDirectory"; ValueData: "{localappdata}\{#ProductName}\ipc-secrets"; Flags: uninsdeletekeyifempty
Root: HKCU; Subkey: "Software\{#ProductName}"; ValueType: string; ValueName: "ExportDirectory"; ValueData: "{localappdata}\{#ProductName}\exports"; Flags: uninsdeletekeyifempty

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
Type: filesandordirs; Name: "{localappdata}\{#ProductName}\logs"
Type: filesandordirs; Name: "{localappdata}\{#ProductName}\ipc-secrets"
Type: filesandordirs; Name: "{localappdata}\{#ProductName}\exports"

