# Installer language selection

The next package includes a dedicated NADABAS language page. Setup itself remains
English. Choices are English (3), Français (4), Português (5), and Bahasa
Indonesia (6). A fresh interactive install requires a choice. Existing valid
preferences are preselected. VBA and the .NET notification UI use the same
per-user registry value. Uninstall preserves this preference.

Silent upgrades preserve the preference. For a fresh silent installation, IT
must supply `/NADABASLANG=3`, `4`, `5`, or `6`. An explicit value also permits an
intentional language change. Invalid values stop Setup before installation.

Smoke packages use only `Software\NADABAS\InstallerSmoke\Language` and must never
alter the real VBA preference. Existing lifecycle, migration and package-update
rehearsals seed French in that isolated profile.

## Acceptance before distributing the next signed package

- Fresh interactive install shows all four native language names and requires a
  choice. Cancel before installation leaves the preference unchanged.
- Select each language in a clean test profile. Start Excel and verify the
  ribbon, About and update notifications use the choice.
- Reinstall interactively and verify the existing language is selected.
- Perform a silent package upgrade and verify the chosen language is unchanged.
- Test fresh silent installation without a language and with an invalid value.
  Both must fail before files or registration change.
- Test explicit silent installation with each supported code.
- Uninstall and verify that the language preference remains.

## Current verification

On 2026-11-05 the complete lifecycle Pascal code, including legacy handoff and
language selection, compiled successfully in an isolated compile-only wrapper.
The Excel bootstrap Release build also passed with no warnings or errors.
The wrapper has no application payload and is not a distributable installer.
Signed installer execution and visual acceptance of this new page are pending.
The previously delivered 0.8.1 installer does not contain this page.
