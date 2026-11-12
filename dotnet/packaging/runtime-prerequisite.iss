// Keep the shared Microsoft runtime serviced independently of NADABAS.
var
  RuntimePage: TWizardPage;

function DesktopRuntimeInstalled: Boolean;
var Names: TArrayOfString; I: Integer;
begin
  Result := False;
  if RegGetValueNames(HKLM32, 'SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App', Names) then
    for I := 0 to GetArrayLength(Names) - 1 do
      if (Copy(Names[I], 1, 5) = '10.0.') and (Pos('-', Names[I]) = 0) then
        Result := True;
end;

procedure OpenRuntimeDownload(Sender: TObject);
var ErrorCode: Integer;
begin
  // Open documentation only. Do not execute downloaded files or elevate Setup.
  if not ShellExecAsOriginalUser('open', 'https://dotnet.microsoft.com/download/dotnet/10.0', '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode) then
    MsgBox('The browser could not be opened. Visit https://dotnet.microsoft.com/download/dotnet/10.0 or ask IT support for .NET 10 Windows Desktop Runtime (x64).', mbError, MB_OK);
end;

procedure InitializeRuntimePage;
var Instructions: TNewStaticText; DownloadButton: TNewButton;
begin
  RuntimePage := CreateCustomPage(wpInfoBefore, 'Microsoft .NET is needed',
    'Install this prerequisite once, then continue NADABAS Setup');
  Instructions := TNewStaticText.Create(RuntimePage);
  Instructions.Parent := RuntimePage.Surface;
  Instructions.Width := RuntimePage.SurfaceWidth;
  Instructions.Height := ScaleY(155);
  Instructions.AutoSize := False;
  Instructions.WordWrap := True;
  Instructions.Caption := 'NADABAS needs Microsoft .NET 10 Windows Desktop Runtime (x64).' + #13#10#13#10 +
    '1. Open the Microsoft download page below.' + #13#10 +
    '2. Under .NET Desktop Runtime 10.0, choose Windows x64 and run the Microsoft installer. The SDK is not needed.' + #13#10 +
    '3. Return here and select Next. Setup will check again.' + #13#10#13#10 +
    'Windows may ask for administrator approval. If you cannot install software, ask IT support to install this prerequisite. Keep NADABAS Setup running under your own Windows account.';
  DownloadButton := TNewButton.Create(RuntimePage);
  DownloadButton.Parent := RuntimePage.Surface;
  DownloadButton.Top := ScaleY(165);
  DownloadButton.Width := ScaleX(240);
  DownloadButton.Height := ScaleY(28);
  DownloadButton.Caption := 'Open Microsoft download page';
  DownloadButton.OnClick := @OpenRuntimeDownload;
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := (PageID = RuntimePage.ID) and DesktopRuntimeInstalled;
end;

function RuntimePageCanContinue(PageID: Integer): Boolean;
begin
  Result := True;
  if (PageID = RuntimePage.ID) and not DesktopRuntimeInstalled then begin
    MsgBox('The required .NET 10 Windows Desktop Runtime (x64) is not installed yet. Finish the Microsoft installation, then select Next again. If a restart is required, restart Windows and run NADABAS Setup again.', mbInformation, MB_OK);
    Result := False;
  end;
end;
