; Isolated manual acceptance package. Never register production Excel add-ins.
#ifndef PayloadDir
  #error PayloadDir is required
#endif
#ifndef UpgradeHash
  #error UpgradeHash is required for the simulated upgrade fixture
#endif
#ifndef ManagedUpgradeHash
  #error ManagedUpgradeHash is required
#endif
#ifndef DatedUpgradeHash
  #error DatedUpgradeHash is required
#endif
#ifndef BetaHash
  #error BetaHash is required
#endif
#ifndef ContextBetaHash
  #error ContextBetaHash is required
#endif
#ifndef DemoHash
  #error DemoHash is required
#endif
#ifndef FixHash
  #error FixHash is required
#endif
#ifdef SmokeTest
  #define PackageId "NadabasLoaderInstallerSmoke"
  #define PackageName "NADABAS Installer Smoke"
  #define ExcelBase "Software\NADABAS\InstallerSmoke\Excel"
#else
  #define PackageId "{{8C18DA56-6F3D-4B42-B114-CC449364E8F9}"
  #define PackageName "NADABAS Loader Test"
  #define ExcelBase "Software\Microsoft\Office\16.0\Excel"
#endif
#ifndef PackageOutput
  #error PackageOutput is required
#endif
#ifndef PackageVersion
  #define PackageVersion "0.9.0"
#endif
[Setup]
AppId={#PackageId}
AppName={#PackageName}
AppVersion={#PackageVersion}
AppPublisher=STATISTISK SENTRALBYRA
DefaultDirName={localappdata}\NADABAS-Loader-Test
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir={#PackageOutput}
OutputBaseFilename=NadabasLoaderTest-{#PackageVersion}-Setup
Compression=lzma2
SolidCompression=yes
CloseApplications=no
RestartApplications=no
SetupLogging=yes
WizardStyle=modern
VersionInfoVersion={#PackageVersion}
VersionInfoDescription=NADABAS manual loader acceptance test
SignTool=nadabas_test
SignedUninstaller=yes
InfoBeforeFile={#PayloadDir}\TESTGUIDE.txt
[Files]
Source: "{#PayloadDir}\*"; DestDir: "{app}"; Excludes: "*.xlam,active.json,store.lock,.nadabas-version-store"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#PayloadDir}\.nadabas-version-store"; DestDir: "{app}"; Flags: onlyifdoesntexist
Source: "{#PayloadDir}\active.json"; DestDir: "{app}"; Flags: onlyifdoesntexist
Source: "{#PayloadDir}\store.lock"; DestDir: "{app}"; Flags: onlyifdoesntexist
Source: "{#PayloadDir}\versions\slot-a\NADABAS.TEST.slot-a.xlam"; DestDir: "{app}\versions\slot-a"; Attribs: readonly; Flags: onlyifdoesntexist
Source: "{#PayloadDir}\versions\slot-b\NADABAS.TEST.slot-b.xlam"; DestDir: "{app}\versions\slot-b"; Attribs: readonly; Flags: onlyifdoesntexist
Source: "{#PayloadDir}\versions\slot-test602006\NADABAS.TEST.slot-test602006.xlam"; DestDir: "{app}\versions\slot-test602006"; Attribs: readonly; Flags: onlyifdoesntexist
Source: "{#PayloadDir}\versions\slot-test602007\NADABAS.TEST.slot-test602007.xlam"; DestDir: "{app}\versions\slot-test602007"; Attribs: readonly; Flags: onlyifdoesntexist
Source: "{#PayloadDir}\versions\slot-test602008\NADABAS.TEST.slot-test602008.xlam"; DestDir: "{app}\versions\slot-test602008"; Attribs: readonly; Flags: onlyifdoesntexist
Source: "{#PayloadDir}\versions\slot-beta602006\NADABAS.TEST.slot-beta602006.xlam"; DestDir: "{app}\versions\slot-beta602006"; Attribs: readonly; Flags: onlyifdoesntexist
Source: "{#PayloadDir}\versions\slot-beta602006b\NADABAS.TEST.slot-beta602006b.xlam"; DestDir: "{app}\versions\slot-beta602006b"; Attribs: readonly; Flags: onlyifdoesntexist
Source: "{#PayloadDir}\versions\slot-demo602006\NADABAS.TEST.slot-demo602006.xlam"; DestDir: "{app}\versions\slot-demo602006"; Attribs: readonly; Flags: onlyifdoesntexist
Source: "{#PayloadDir}\versions\slot-fix602009\NADABAS.TEST.slot-fix602009.xlam"; DestDir: "{app}\versions\slot-fix602009"; Attribs: readonly; Flags: onlyifdoesntexist
[Icons]
Name: "{userprograms}\{#PackageName}\Test guide"; Filename: "{app}\TESTGUIDE.txt"
Name: "{userprograms}\{#PackageName}\Open installation folder"; Filename: "{app}"
[UninstallDelete]
Type: files; Name: "{app}\download.lock"
Type: files; Name: "{app}\last-download-check.json"
Type: files; Name: "{app}\skipped-package-version"
Type: files; Name: "{app}\pending-platform-update"
Type: files; Name: "{app}\platform-update.lock"
Type: files; Name: "{app}\legacy-migration.required"
Type: files; Name: "{app}\pending-update.json"
[Run]
Filename: "{app}\TESTGUIDE.txt"; Description: "Open the manual test guide"; Flags: shellexec postinstall skipifsilent nowait
[Code]
#include "loader-lifecycle.iss"
