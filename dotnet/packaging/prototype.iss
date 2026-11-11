; TEST ONLY. Same signed NADABAS 6.02.005 payload in each package version.
; No Excel registration or startup hooks. No application/process termination.
#ifndef PackageVersion
  #error PackageVersion must be supplied
#endif
#ifndef PayloadDir
  #error PayloadDir must be supplied
#endif
#ifndef PackageOutput
  #error PackageOutput must be supplied
#endif
[Setup]
AppId={{20788D4E-ED0E-4DBC-A2B2-3043CE45C9D7}
AppName=NADABAS Installation Prototype (TEST ONLY)
AppVersion={#PackageVersion}
AppPublisher=NADABAS development
DefaultDirName={localappdata}\NADABAS-Installation-Prototype
PrivilegesRequired=lowest
DisableDirPage=yes
DisableProgramGroupPage=yes
UninstallDisplayName=NADABAS Installation Prototype (TEST ONLY)
OutputDir={#PackageOutput}
OutputBaseFilename=NadabasPrototype-{#PackageVersion}-Setup
Compression=lzma2
SolidCompression=yes
CloseApplications=no
RestartApplications=no
Uninstallable=yes
SetupLogging=yes
WizardStyle=modern
VersionInfoVersion={#PackageVersion}
VersionInfoDescription=NADABAS installer experiment - not a release
[Files]
Source: "{#PayloadDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
