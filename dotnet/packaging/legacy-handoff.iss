// Only the supported per-user 6.02.005 path and a single REG_SZ startup entry
// qualify. The journal and old file remain available after uninstall.
function ReadValueInfo(Key: Cardinal; SubKey, Name: String; Flags: Cardinal;
  var Kind: Cardinal; Data: Integer; var Bytes: Cardinal): Integer;
  external 'RegGetValueW@advapi32.dll stdcall';

function PlainString(const Name: String): Boolean;
var Kind, Bytes: Cardinal;
begin
  Bytes := 0;
  Result := (ReadValueInfo($80000001, OptionsKey, Name, $10010002, Kind, 0, Bytes) = 0) and (Kind = 1);
end;

function LocalDriveType(Path: String): Cardinal;
  external 'GetDriveTypeW@kernel32.dll stdcall';

function OldAddin: String;
begin
#ifdef SmokeTest
  Result := ExtractFileDir(ExpandConstant('{app}')) + '\profile\AddIns\NADABAS.XLAM';
#else
  Result := ExpandConstant('{userappdata}\Microsoft\AddIns\NADABAS.XLAM');
#endif
end;

function BackupDir: String;
begin
  Result := ExpandConstant('{app}') + '-legacy-backup';
end;

function LoaderEntry: String;
begin
  Result := '/R "' + ExpandConstant('{app}\Nadabas.Loader.Test.xll') + '"';
end;

procedure CheckLocalPath(Path: String);
var Parent: String; Attr: Cardinal;
begin
  if (Length(Path) < 3) or (Copy(Path, 2, 2) <> ':\') or
    (LocalDriveType(ExtractFileDrive(Path) + '\') <> 3) then
    RaiseException('Legacy migration requires a local fixed drive.');
  while Length(Path) > 3 do begin
    Attr := FileAttributes(Path);
    if (Attr <> $FFFFFFFF) and ((Attr and $400) <> 0) then
      RaiseException('Legacy migration requires a path without filesystem links.');
    Parent := ExtractFileDir(Path);
    if Parent = Path then Break;
    Path := Parent;
  end;
end;

procedure RequireBaseline(const Path: String);
begin
  CheckLocalPath(Path);
  if not FileExists(Path) then RaiseException('Missing migration file: ' + Path);
  if CompareText(GetSHA256OfFile(Path), BaselineHash) <> 0 then
    RaiseException('Legacy file differs from approved 6.02.005. Review it before continuing.');
end;

procedure ReadBackup(var Slot, Original: String);
var Lines: TArrayOfString; Path: String;
begin
  Path := BackupDir + '\registration.txt';
  CheckLocalPath(Path);
  RequireBaseline(BackupDir + '\NADABAS.XLAM');
  if not LoadStringsFromFile(Path, Lines) then RaiseException('Cannot read legacy registration backup.');
  if GetArrayLength(Lines) <> 3 then RaiseException('Invalid legacy registration backup.');
  if (Lines[0] <> 'NADABAS legacy registration v1 REG_SZ') or not IsOpenName(Lines[1]) or
    not RefersTo(Lines[2], OldAddin) then RaiseException('Unrecognized legacy registration backup.');
  Slot := Lines[1]; Original := Lines[2];
end;

function OurPayload(const Value: String): Boolean;
var Slots: TArrayOfString; I: Integer; Path: String;
begin
  Result := False;
  SetArrayLength(Slots, 9);
  Slots[0] := 'slot-a';
  Slots[1] := 'slot-b';
  Slots[2] := 'slot-test602006';
  Slots[3] := 'slot-test602007';
  Slots[4] := 'slot-test602008';
  Slots[5] := 'slot-beta602006';
  Slots[6] := 'slot-beta602006b';
  Slots[7] := 'slot-demo602006';
  Slots[8] := 'slot-fix602009';
  for I := 0 to GetArrayLength(Slots) - 1 do begin
    Path := ExpandConstant('{app}') + '\versions\' + Slots[I] + '\NADABAS.TEST.' + Slots[I] + '.xlam';
    if RefersTo(Value, Path) then begin Result := True; Exit; end;
  end;
end;

procedure FindLegacy(var Slot, Original: String);
var Names: TArrayOfString; I, Loaders: Integer; Value: String;
begin
  Slot := ''; Original := ''; Loaders := 0;
  if not RegKeyExists(HKCU64, OptionsKey) then Exit;
  if not RegGetValueNames(HKCU64, OptionsKey, Names) then RaiseException('Cannot inspect Excel startup entries.');
  for I := 0 to GetArrayLength(Names) - 1 do begin
    if IsOpenName(Names[I]) and RegQueryStringValue(HKCU64, OptionsKey, Names[I], Value) then begin
      if RefersTo(Value, OldAddin) then begin
        if (Slot <> '') or not PlainString(Names[I]) then RaiseException('Ambiguous legacy Excel startup registration.');
        Slot := Names[I]; Original := Value;
      end else if RefersTo(Value, ExpandConstant('{app}\Nadabas.Loader.Test.xll')) then
        Loaders := Loaders + 1
      else if not OurPayload(Value) and (Pos('NADABAS', UpperCase(Value)) > 0) and
        ((Pos('.XLA', UpperCase(Value)) > 0) or (Pos('.XLL', UpperCase(Value)) > 0)) then
        RaiseException('A custom NADABAS startup entry requires review.');
    end;
  end;
  if (Loaders > 1) or ((Slot <> '') and (Loaders <> 0)) then RaiseException('Multiple NADABAS startup entries require review.');
end;

procedure RequireEntry(const Slot, Expected: String);
var Value: String;
begin
  if not PlainString(Slot) or not RegQueryStringValue(HKCU64, OptionsKey, Slot, Value) or
    (Value <> Expected) then RaiseException('The Excel startup registration changed. Review it before continuing.');
end;

procedure CheckRestore(var Slot, Original: String);
var Value: String;
begin
  ReadBackup(Slot, Original);
  if not PlainString(Slot) or not RegQueryStringValue(HKCU64, OptionsKey, Slot, Value) or
    ((Value <> LoaderEntry) and (Value <> Original)) then RaiseException('Legacy rollback conflicts with the current startup registration.');
  CheckLocalPath(OldAddin);
  if FileExists(OldAddin) then RequireBaseline(OldAddin);
end;

procedure ReviewStartupFolder(const Path: String);
var Item: TFindRec; Ext: String;
begin
  if Path = '' then Exit;
  CheckLocalPath(Path);
  if FindFirst(AddBackslash(Path) + '*', Item) then begin
    try
      repeat
        Ext := UpperCase(ExtractFileExt(Item.Name));
        if (Pos('NADABAS', UpperCase(Item.Name)) > 0) and
          ((Ext = '.XLA') or (Ext = '.XLAM') or (Ext = '.XLSM') or (Ext = '.XLL')) then
          RaiseException('Review the NADABAS file in the Excel startup folder before installing.');
      until not FindNext(Item);
    finally FindClose(Item); end;
  end;
end;

function LegacyPreflight: String;
var Slot, Original, StoredSlot, StoredOriginal, Path: String;
begin
  Result := '';
  try
#ifdef SmokeTest
    ReviewStartupFolder(ExtractFileDir(ExpandConstant('{app}')) + '\profile\XLSTART');
#else
    ReviewStartupFolder(ExpandConstant('{userappdata}\Microsoft\Excel\XLSTART'));
    if RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', '', Path) then
      ReviewStartupFolder(ExtractFileDir(Path) + '\XLSTART');
#endif
    if RegQueryStringValue(HKCU64, OptionsKey, 'AltStartup', Path) then ReviewStartupFolder(Path);
    FindLegacy(Slot, Original);
    CheckLocalPath(BackupDir);
    if DirExists(BackupDir) then begin
      CheckRestore(StoredSlot, StoredOriginal);
      if (Slot <> '') and ((Slot <> StoredSlot) or (Original <> StoredOriginal)) then
        RaiseException('The legacy registration differs from the saved migration.');
    end else if FileExists(ExpandConstant('{app}\legacy-migration.required')) then
      RaiseException('Restore the missing legacy migration backup before continuing.');
    if Slot <> '' then RequireBaseline(OldAddin);
  except
    Result := GetExceptionMessage;
  end;
end;

function SwitchLegacy: Boolean;
var Slot, Original, Staging, CurrentSlot, CurrentValue: String; Lines: TArrayOfString;
begin
  Result := False;
  FindLegacy(Slot, Original);
  if Slot = '' then Exit;
  RequireBaseline(OldAddin);
  if not DirExists(BackupDir) then begin
    Staging := BackupDir + '.preparing';
    CheckLocalPath(Staging);
    if not CreateDir(Staging) then RaiseException('Cannot create a new migration backup.');
    if not CopyFile(OldAddin, Staging + '\NADABAS.XLAM', True) then RaiseException('Cannot back up the old add-in.');
    RequireBaseline(Staging + '\NADABAS.XLAM');
    SetArrayLength(Lines, 3);
    Lines[0] := 'NADABAS legacy registration v1 REG_SZ'; Lines[1] := Slot; Lines[2] := Original;
    if not SaveStringsToUTF8File(Staging + '\registration.txt', Lines, False) then RaiseException('Cannot save migration metadata.');
    if not RenameFile(Staging, BackupDir) then RaiseException('Cannot commit the legacy backup.');
  end;
  ReadBackup(CurrentSlot, CurrentValue);
  if (CurrentSlot <> Slot) or (CurrentValue <> Original) then RaiseException('Legacy backup does not match this registration.');
  if not SaveStringToFile(ExpandConstant('{app}\legacy-migration.required'), '1', False) then RaiseException('Cannot record legacy migration.');
  if not ExcelClosed then RaiseException('Close Excel and retry migration.');
  RequireBaseline(OldAddin);
  RequireEntry(Slot, Original);
  if not RegWriteStringValue(HKCU64, OptionsKey, Slot, LoaderEntry) then RaiseException('Cannot switch Excel startup to the loader.');
  Result := True;
end;

procedure RestoreLegacy;
var Slot, Original: String;
begin
  if not DirExists(BackupDir) and not FileExists(ExpandConstant('{app}\legacy-migration.required')) then Exit;
  CheckRestore(Slot, Original);
  if not FileExists(OldAddin) then begin
    if not CopyFile(BackupDir + '\NADABAS.XLAM', OldAddin, True) then RaiseException('Cannot restore the legacy add-in.');
    RequireBaseline(OldAddin);
  end;
  CheckRestore(Slot, Original);
  if not RegWriteStringValue(HKCU64, OptionsKey, Slot, Original) then RaiseException('Cannot restore legacy startup registration.');
end;
