// Compiled into Setup and Uninstall. No Office trust/security settings are changed.
const
  OptionsKey = '{#ExcelBase}\Options';
  ManagerKey = '{#ExcelBase}\Add-in Manager';
  BaselineHash = '2F7432B9B04C414337EB1AFEB07E2B93F05E4A5A6D151073FA967A2F7EA89A85';
  UpgradeHash = '{#UpgradeHash}';
  ManagedUpgradeHash = '{#ManagedUpgradeHash}';
  DatedUpgradeHash = '{#DatedUpgradeHash}';
  BetaHash = '{#BetaHash}';
  ContextBetaHash = '{#ContextBetaHash}';
  DemoHash = '{#DemoHash}';
  FixHash = '{#FixHash}';

function FileAttributes(Name: String): Cardinal;
  external 'GetFileAttributesW@kernel32.dll stdcall';
function BinaryType(Name: String; var Kind: Cardinal): Boolean;
  external 'GetBinaryTypeW@kernel32.dll stdcall';

function PrerequisiteError: String;
var Names: TArrayOfString; I: Integer; RuntimeFound: Boolean; ExcelPath: String; Kind: Cardinal;
begin
  Result := '';
  RuntimeFound := False;
  if RegGetValueNames(HKLM32, 'SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App', Names) then
    for I := 0 to GetArrayLength(Names) - 1 do
      if (Copy(Names[I], 1, 5) = '10.0.') and (Pos('-', Names[I]) = 0) then RuntimeFound := True;
  if not RuntimeFound then begin
    Result := 'Microsoft .NET 10 Windows Desktop Runtime (x64) is required. Install it from https://dotnet.microsoft.com/download/dotnet/10.0 and run Setup again. The SDK is not required.';
    Exit;
  end;
  if not RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', '', ExcelPath) then begin
    Result := 'Microsoft Excel (64-bit) is required. No Excel installation was found.';
    Exit;
  end;
  if not BinaryType(ExcelPath, Kind) then begin Result := 'Unable to verify the Excel installation architecture.'; Exit; end;
  if Kind <> 6 then Result := 'This test package supports only 64-bit Excel.';
end;

function ExcelClosed: Boolean;
var Locator, Service, Processes: Variant;
begin
  Result := False;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    Service := Locator.ConnectServer('', 'root\CIMV2');
    Processes := Service.ExecQuery('SELECT ProcessId FROM Win32_Process WHERE Name = ''EXCEL.EXE''');
    Result := Processes.Count = 0;
  except
    Log('Unable to verify that Excel is closed; refusing to modify installation.');
  end;
end;

function SafeTree(const Path: String; ResidualOnly, InSessions: Boolean): Boolean;
var Item: TFindRec;
begin
  Result := False;
  if (FileAttributes(Path) and $400) <> 0 then Exit;
  Result := True;
  if FindFirst(AddBackslash(Path) + '*', Item) then begin
    try
      repeat
        if (Item.Name <> '.') and (Item.Name <> '..') then begin
          if (Item.Attributes and $400) <> 0 then begin Result := False; Exit; end;
          if (Item.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0 then begin
            if ResidualOnly and ((Item.Name <> 'sessions') or InSessions) then begin Result := False; Exit; end;
            if not SafeTree(AddBackslash(Path) + Item.Name, ResidualOnly, True) then begin Result := False; Exit; end;
          end else if ResidualOnly then begin
            if not ((InSessions and (CompareText(ExtractFileExt(Item.Name), '.json') = 0)) or
              ((not InSessions) and (Item.Name = 'store.lock'))) then begin Result := False; Exit; end;
          end;
        end;
      until not FindNext(Item);
    finally FindClose(Item); end;
  end;
end;

function SafeResidual(const Path, Relative: String): Boolean;
var Item: TFindRec; Child, Rel: String;
begin
  Result := True;
  if FindFirst(AddBackslash(Path) + '*', Item) then begin
    try
      repeat
        if (Item.Name <> '.') and (Item.Name <> '..') then begin
          Child := AddBackslash(Path) + Item.Name;
          Rel := Relative + Item.Name;
          if (Item.Attributes and $400) <> 0 then begin Result := False; Exit; end;
          if (Item.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0 then begin
            if (Rel <> 'sessions') and (Rel <> 'tools') and (Rel <> 'updater') and (Rel <> 'versions') and
              (Rel <> 'versions\slot-a') and (Rel <> 'versions\slot-b') and
              (Rel <> 'versions\slot-test602006') and (Rel <> 'versions\slot-test602007') and (Rel <> 'versions\slot-test602008') and (Rel <> 'versions\slot-beta602006') and (Rel <> 'versions\slot-beta602006b') and (Rel <> 'versions\slot-demo602006') and (Rel <> 'versions\slot-fix602009') then begin Result := False; Exit; end;
            if not SafeResidual(Child, Rel + '\') then begin Result := False; Exit; end;
          end else if (Relative = 'sessions\') and (CompareText(ExtractFileExt(Item.Name), '.json') = 0) then begin
            // Preserve historical reports.
          end else if (Rel = 'store.lock') or (Rel = 'pending-update.json') then begin
            // The version-store command may recreate this lock file.
          end else if (Rel = 'versions\slot-a\NADABAS.TEST.slot-a.xlam') or
            (Rel = 'versions\slot-b\NADABAS.TEST.slot-b.xlam') then begin
            if CompareText(GetSHA256OfFile(Child), BaselineHash) <> 0 then begin Result := False; Exit; end;
          end else if Rel = 'versions\slot-test602006\NADABAS.TEST.slot-test602006.xlam' then begin
            if CompareText(GetSHA256OfFile(Child), UpgradeHash) <> 0 then begin Result := False; Exit; end;
          end else if Rel = 'versions\slot-test602007\NADABAS.TEST.slot-test602007.xlam' then begin
            if CompareText(GetSHA256OfFile(Child), ManagedUpgradeHash) <> 0 then begin Result := False; Exit; end;
          end else if Rel = 'versions\slot-test602008\NADABAS.TEST.slot-test602008.xlam' then begin
            if CompareText(GetSHA256OfFile(Child), DatedUpgradeHash) <> 0 then begin Result := False; Exit; end;
          end else if Rel = 'versions\slot-beta602006\NADABAS.TEST.slot-beta602006.xlam' then begin
            if CompareText(GetSHA256OfFile(Child), BetaHash) <> 0 then begin Result := False; Exit; end;
          end else if Rel = 'versions\slot-beta602006b\NADABAS.TEST.slot-beta602006b.xlam' then begin
            if CompareText(GetSHA256OfFile(Child), ContextBetaHash) <> 0 then begin Result := False; Exit; end;
          end else if Rel = 'versions\slot-demo602006\NADABAS.TEST.slot-demo602006.xlam' then begin
            if CompareText(GetSHA256OfFile(Child), DemoHash) <> 0 then begin Result := False; Exit; end;
          end else if Rel = 'versions\slot-fix602009\NADABAS.TEST.slot-fix602009.xlam' then begin
            if CompareText(GetSHA256OfFile(Child), FixHash) <> 0 then begin Result := False; Exit; end;
          end else begin Result := False; Exit; end;
        end;
      until not FindNext(Item);
    finally FindClose(Item); end;
  end;
end;

function IsOpenName(const Name: String): Boolean;
var I: Integer;
begin
  Result := False;
  if CompareText(Copy(Name, 1, 4), 'OPEN') <> 0 then Exit;
  for I := 5 to Length(Name) do
    if (Name[I] < '0') or (Name[I] > '9') then Exit;
  Result := True;
end;

function RefersTo(const Value, Path: String): Boolean;
begin
  Result := (CompareText(Value, Path) = 0) or
    (CompareText(Value, '"' + Path + '"') = 0) or
    (CompareText(Value, '/R "' + Path + '"') = 0);
end;

#include "legacy-handoff.iss"
#include "language-selection.iss"

function PrepareToInstall(var NeedsRestart: Boolean): String;
var Root, Ancestor, Slot, Path, ExpectedHash: String; Lines: TArrayOfString; I: Integer;
begin
  Result := LanguagePreflight;
  if Result <> '' then Exit;
  Result := PrerequisiteError;
  if Result <> '' then Exit;
  if not ExcelClosed then begin Result := 'Save your work and close all Excel windows, then retry. Excel will not be closed automatically.'; Exit; end;
  Root := ExpandConstant('{app}');
  Ancestor := Root;
  while Length(Ancestor) > 3 do begin
    if DirExists(Ancestor) then
      if (FileAttributes(Ancestor) and $400) <> 0 then begin Result := 'Installation through filesystem links is not supported.'; Exit; end;
    Ancestor := ExtractFileDir(Ancestor);
  end;
  Result := LegacyPreflight;
  if Result <> '' then Exit;
  if not DirExists(Root) then Exit;
  if FileExists(AddBackslash(Root) + '.nadabas-version-store') then begin
    if not SafeTree(Root, False, False) then begin Result := 'The existing installation contains a filesystem link. No files were changed.'; Exit; end;
    if not LoadStringsFromFile(AddBackslash(Root) + '.nadabas-version-store', Lines) then begin Result := 'Cannot read the installation marker.'; Exit; end;
    if GetArrayLength(Lines) <> 1 then begin Result := 'Unrecognized installation directory.'; Exit; end;
    if Lines[0] <> 'NADABAS version-store prototype v1' then begin Result := 'Unrecognized installation directory.'; Exit; end;
    for I := 0 to 8 do begin
      ExpectedHash := BaselineHash;
      if I = 0 then Slot := 'slot-a'
      else if I = 1 then Slot := 'slot-b'
      else if I = 2 then begin Slot := 'slot-test602006'; ExpectedHash := UpgradeHash; end
      else if I = 3 then begin Slot := 'slot-test602007'; ExpectedHash := ManagedUpgradeHash; end
      else if I = 4 then begin Slot := 'slot-test602008'; ExpectedHash := DatedUpgradeHash; end
      else if I = 5 then begin Slot := 'slot-beta602006'; ExpectedHash := BetaHash; end
      else if I = 6 then begin Slot := 'slot-beta602006b'; ExpectedHash := ContextBetaHash; end
      else if I = 7 then begin Slot := 'slot-demo602006'; ExpectedHash := DemoHash; end
      else begin Slot := 'slot-fix602009'; ExpectedHash := FixHash; end;
      Path := AddBackslash(Root) + 'versions\' + Slot + '\NADABAS.TEST.' + Slot + '.xlam';
      if FileExists(Path) then
        if CompareText(GetSHA256OfFile(Path), ExpectedHash) <> 0 then begin Result := 'An existing test version has changed. Installation stopped and preserved it.'; Exit; end;
    end;
    Log('Updating known test installation; preserving active.json, version slots and reports.');
  end else if not SafeResidual(Root, '') then
    Result := 'The destination contains unrecognized files. They have been left unchanged.'
  else Log('Reusing empty directory or session reports left after uninstall.');
end;

procedure RemoveExactRegistration(const Path: String);
var Names: TArrayOfString; I: Integer; Value: String;
begin
  if RegGetValueNames(HKCU64, OptionsKey, Names) then
    for I := 0 to GetArrayLength(Names) - 1 do
      if IsOpenName(Names[I]) then
        if RegQueryStringValue(HKCU64, OptionsKey, Names[I], Value) then
          if RefersTo(Value, Path) then
            if not RegDeleteValue(HKCU64, OptionsKey, Names[I]) then RaiseException('Cannot remove test startup entry.');
  if RegValueExists(HKCU64, ManagerKey, Path) then
    if not RegDeleteValue(HKCU64, ManagerKey, Path) then RaiseException('Cannot remove test add-in entry.');
end;

procedure CleanupPayloadRegistrations;
begin
  // Exact paths from our older prototype only; unrelated NADABAS installations remain untouched.
  RemoveExactRegistration(ExpandConstant('{app}\versions\slot-a\NADABAS.TEST.slot-a.xlam'));
  RemoveExactRegistration(ExpandConstant('{app}\versions\slot-b\NADABAS.TEST.slot-b.xlam'));
  RemoveExactRegistration(ExpandConstant('{app}\versions\slot-test602006\NADABAS.TEST.slot-test602006.xlam'));
  RemoveExactRegistration(ExpandConstant('{app}\versions\slot-test602007\NADABAS.TEST.slot-test602007.xlam'));
  RemoveExactRegistration(ExpandConstant('{app}\versions\slot-test602008\NADABAS.TEST.slot-test602008.xlam'));
  RemoveExactRegistration(ExpandConstant('{app}\versions\slot-beta602006\NADABAS.TEST.slot-beta602006.xlam'));
  RemoveExactRegistration(ExpandConstant('{app}\versions\slot-beta602006b\NADABAS.TEST.slot-beta602006b.xlam'));
  RemoveExactRegistration(ExpandConstant('{app}\versions\slot-demo602006\NADABAS.TEST.slot-demo602006.xlam'));
  RemoveExactRegistration(ExpandConstant('{app}\versions\slot-fix602009\NADABAS.TEST.slot-fix602009.xlam'));
end;

procedure CurStepChanged(CurStep: TSetupStep);
var Name, Path, Value: String; I, ExitCode: Integer; Names: TArrayOfString; Found: Boolean;
begin
  if CurStep = ssInstall then
    if not ExcelClosed then RaiseException('Excel was opened during setup. Close Excel and retry.');
  if CurStep <> ssPostInstall then Exit;
  if not ExcelClosed then RaiseException('Excel was opened during setup. Close Excel and run setup again to finish registration.');
  // Explicit test-release upgrade: validate bytes through the compiled catalog
  // and atomically select the new slot only after payload installation succeeds.
  if not Exec(ExpandConstant('{app}\tools\Nadabas.VersionStore.exe'),
    'activate "' + ExpandConstant('{app}') + '" slot-fix602009', '', SW_HIDE, ewWaitUntilTerminated, ExitCode) then
    RaiseException('Unable to start release activation. Previous selection has not been intentionally changed.');
  if ExitCode <> 0 then RaiseException('Release verification or activation failed. See the installation log.');
  Path := ExpandConstant('{app}\Nadabas.Loader.Test.xll');
  CleanupPayloadRegistrations;
  if SwitchLegacy then begin SaveSelectedLanguage; Exit; end;
  Found := False;
  if RegGetValueNames(HKCU64, OptionsKey, Names) then
    for I := 0 to GetArrayLength(Names) - 1 do
      if IsOpenName(Names[I]) then
        if RegQueryStringValue(HKCU64, OptionsKey, Names[I], Value) then
          if RefersTo(Value, Path) then Found := True;
  if not Found then begin
    I := 0;
    repeat
      if I = 0 then Name := 'OPEN' else Name := 'OPEN' + IntToStr(I);
      I := I + 1;
    until not RegValueExists(HKCU64, OptionsKey, Name);
    if not RegWriteStringValue(HKCU64, OptionsKey, Name, '/R "' + Path + '"') then
      RaiseException('Cannot register the test loader for Excel startup.');
  end;
  Log('Test loader registered for next Excel startup.');
  SaveSelectedLanguage;
end;

function InitializeUninstall: Boolean;
var Slot, Original: String;
begin
  Result := ExcelClosed;
  if Result and (DirExists(BackupDir) or FileExists(ExpandConstant('{app}\legacy-migration.required'))) then begin
    try CheckRestore(Slot, Original); except Log(GetExceptionMessage); SuppressibleMsgBox(GetExceptionMessage, mbError, MB_OK, IDOK); Result := False; Exit; end;
  end;
  if not Result then SuppressibleMsgBox('Save your work and close Excel before uninstalling.', mbError, MB_OK, IDOK);
end;

procedure CurUninstallStepChanged(CurStep: TUninstallStep);
begin
  if CurStep <> usUninstall then Exit;
  if not ExcelClosed then RaiseException('Close Excel before uninstalling.');
  RestoreLegacy;
  RemoveExactRegistration(ExpandConstant('{app}\Nadabas.Loader.Test.xll'));
  CleanupPayloadRegistrations;
end;
