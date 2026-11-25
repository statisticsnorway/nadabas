// Native Setup discovery: no modern .NET installation is needed to select a payload.
var
  SelectedArchitecture: AnsiString;

function PeArchitecture(const Path: String): String;
var Data: AnsiString; Offset: Integer; Machine: Integer;
begin
  Result := '';
  if not LoadStringFromFile(Path, Data) then Exit;
  if (Length(Data) < 64) or (Copy(Data, 1, 2) <> 'MZ') then Exit;
  // Reject large/negative offsets before indexing the PE header.
  if (Ord(Data[63]) <> 0) or (Ord(Data[64]) <> 0) then Exit;
  Offset := Ord(Data[61]) + 256 * Ord(Data[62]);
  if (Offset < 64) or (Offset + 26 > Length(Data)) then Exit;
  if Copy(Data, Offset + 1, 4) <> 'PE' + #0 + #0 then Exit;
  Machine := Ord(Data[Offset + 5]) + 256 * Ord(Data[Offset + 6]);
  if (Machine = $14C) and (Ord(Data[Offset + 25]) = $0B) and (Ord(Data[Offset + 26]) = 1) then Result := 'x86';
  if (Machine = $8664) and (Ord(Data[Offset + 25]) = $0B) and (Ord(Data[Offset + 26]) = 2) then Result := 'x64';
end;

function PrivatePrerequisiteError: String;
var Path, Candidate: String; Existing: AnsiString; Release: Cardinal; View: Integer;
  Item: TFindRec;
begin
  Result := '';
  SelectedArchitecture := '';
#ifdef SmokeTest
  SelectedArchitecture := ExpandConstant('{param:TESTEXCELARCH|}');
  if (SelectedArchitecture <> '') and (SelectedArchitecture <> 'x86') and (SelectedArchitecture <> 'x64') then begin
    Result := 'Invalid isolated test architecture.'; Exit;
  end;
#endif
  if SelectedArchitecture = '' then
    for View := 0 to 1 do begin
      if (View = 0) or IsWin64 then begin
        if View = 0 then Release := HKLM32 else Release := HKLM64;
        if RegQueryStringValue(Release, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', '', Path) then begin
          Candidate := PeArchitecture(Path);
          if Candidate = '' then begin Result := 'Unable to verify the registered Excel executable. Ask IT to repair the Office registration.'; Exit; end;
          if (SelectedArchitecture <> '') and (SelectedArchitecture <> Candidate) then begin
            Result := 'Conflicting Excel architectures were found. Ask IT to review the Office installations.'; Exit;
          end;
          SelectedArchitecture := Candidate;
        end;
      end;
    end;
  if SelectedArchitecture = '' then begin Result := 'Microsoft Excel (32-bit or 64-bit) is required.'; Exit; end;
  if (SelectedArchitecture = 'x64') and not IsWin64 then begin Result := '64-bit Excel requires 64-bit Windows.'; Exit; end;
  if not RegQueryDWordValue(HKLM32, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release) then Release := 0;
  if Release < 528040 then begin Result := 'Microsoft .NET Framework 4.8 is required. Ask IT to enable it. Modern .NET is included with NADABAS.'; Exit; end;
  if FileExists(ExpandConstant('{app}\excel-architecture.txt')) then begin
    if not LoadStringFromFile(ExpandConstant('{app}\excel-architecture.txt'), Existing) then begin
      Result := 'Cannot read the installed Excel architecture.'; Exit;
    end;
    if Trim(Existing) <> SelectedArchitecture then
      Result := 'Excel architecture has changed. Uninstall NADABAS before installing it for the new Excel architecture. Your workbooks and databases are not removed.';
  end;
  if Result <> '' then Exit;
  // Never leave two runtime patches behind: startup verifies an exact inventory.
  // Until a reviewed runtime-retirement transaction exists, refuse before writes.
  for View := 0 to 2 do begin
    if View = 0 then Candidate := 'host\fxr'
    else if View = 1 then Candidate := 'shared\Microsoft.NETCore.App'
    else Candidate := 'shared\Microsoft.WindowsDesktop.App';
    Path := ExpandConstant('{app}\runtime\') + SelectedArchitecture + '\' + Candidate;
    if FindFirst(Path + '\*', Item) then begin
      try
        repeat
          if (Item.Name <> '.') and (Item.Name <> '..') and
            ((Item.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0) and
            (Item.Name <> '{#PrivateRuntimeVersion}') then begin
            Result := 'This update changes the private .NET runtime. Keep the current installation and contact the release maintainer for a reviewed runtime upgrade.';
            Exit;
          end;
        until not FindNext(Item);
      finally FindClose(Item); end;
    end;
  end;
end;

function InstallX86: Boolean;
begin Result := SelectedArchitecture = 'x86'; end;

function InstallX64: Boolean;
begin Result := SelectedArchitecture = 'x64'; end;
