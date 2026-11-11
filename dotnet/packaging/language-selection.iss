// Use the same per-user setting as VBA and the .NET update notifications.
// Smoke builds must never change the user's real NADABAS language.
const
#ifdef SmokeTest
  LanguageKey = 'Software\NADABAS\InstallerSmoke\Language';
#else
  LanguageKey = 'Software\VB and VBA Program Settings\NADABAS\Language';
#endif
var
  LanguagePage: TInputOptionWizardPage;
  RequestedLanguage: String;

function LanguageIndex(const Code: String): Integer;
begin
  Result := -1;
  if Code = '3' then Result := 0
  else if Code = '4' then Result := 1
  else if Code = '5' then Result := 2
  else if Code = '6' then Result := 3;
end;

procedure InitializeWizard;
var Existing: String; Index: Integer;
begin
  RequestedLanguage := ExpandConstant('{param:NADABASLANG|}');
  LanguagePage := CreateInputOptionPage(wpInfoBefore, 'NADABAS language',
    'Choose the language to use in NADABAS',
    'Select your language for Excel and NADABAS update messages. You can change it later in NADABAS Settings > Language.', True, False);
  LanguagePage.Add('English');
  LanguagePage.Add('Français');
  LanguagePage.Add('Português');
  LanguagePage.Add('Bahasa Indonesia');
  Index := -1;
  if RegQueryStringValue(HKCU, LanguageKey, 'Code', Existing) then
    Index := LanguageIndex(Existing);
  if RequestedLanguage <> '' then Index := LanguageIndex(RequestedLanguage);
  LanguagePage.SelectedValueIndex := Index;
end;

function LanguagePreflight: String;
begin
  Result := '';
  if (RequestedLanguage <> '') and (LanguageIndex(RequestedLanguage) < 0) then
    Result := 'Invalid /NADABASLANG value. Use 3 (English), 4 (French), 5 (Portuguese) or 6 (Bahasa Indonesia).';
  if (Result = '') and WizardSilent and (LanguagePage.SelectedValueIndex < 0) then
    Result := 'For a first silent installation, specify /NADABASLANG=3, 4, 5 or 6. Interactive Setup lets the user choose.';
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if CurPageID = LanguagePage.ID then
    if LanguagePage.SelectedValueIndex < 0 then begin
      MsgBox('Please select your NADABAS language to continue.', mbInformation, MB_OK);
      Result := False;
    end;
end;

procedure SaveSelectedLanguage;
var Code: String;
begin
  // Silent updates preserve the existing preference. Fresh silent installs
  // require an explicit language because older XLAMs may skip first-run prompts.
  if WizardSilent and (RequestedLanguage = '') then Exit;
  if LanguagePage.SelectedValueIndex < 0 then Exit;
  Code := IntToStr(LanguagePage.SelectedValueIndex + 3);
  if not RegWriteStringValue(HKCU, LanguageKey, 'Code', Code) then
    RaiseException('Unable to save your NADABAS language preference.');
end;
