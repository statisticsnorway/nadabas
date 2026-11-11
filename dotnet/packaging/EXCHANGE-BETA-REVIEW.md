# Exchange beta review and package test, 2026-11-03

Input: user-supplied NADABAS_BETA.zip, SHA-256
855DB9C8DE23953340F7EB867A01187BF29877FE339DDBABAD8BCD018C7F6C41.
Contains NADABAS.xlam and NADABAS.6.02.006.xlsm. Original XLAM is unsigned,
reports version 6.02.006, and differs from the earlier simulated 006 fixture.
Original ZIP and extracted review files are preserved; no macros ran during review.
The old XLSM installer is not part of the new package.

Both beta ribbon XMLs match the repository's original ribbon XMLs. Five Exchange
groups are controlled by three callbacks. Those callbacks depend on ExchLinks
and administrator status, with no independent user display preference. This
supports adding a display gate; it does not prove the only cause of the reported
compressed ribbon. Existing cached workbook flags and ribbon refresh behavior
may need further investigation if problems persist after the display change.

User chose to fix the display before testing the package. Changes:
- Show Exchange tools toggle remains available beside About, even with no database.
- Per-user HKCU VBA settings preference defaults off; callback failures also hide.
- All three visibility callbacks preserve existing workbook/admin checks.
- Toggle invalidates the ribbon immediately and persists across restarts.
- No Exchange connection, link or data changes; this is not a server-disable switch.
- Candidate adds the existing managed update bridge and fixed build date.

Macro-disabled comparison of the signed candidate against the original found
only RibbonUI, InterfaceVersionUpdate and dlgAbout code changed. Authenticode and
VBASigned checks passed. Both XML schemas carry the same toggle. Source tests
cover callback mapping, default state and narrow transformation validation.
Live Excel rendering, persistence and Exchange operations still need user tests.

Package 0.7.1 uses slot-beta602006 and preserves version 6.02.006. Installing it
over 0.7.0 is an explicit beta selection, not a numeric VBA upgrade from simulated
6.02.008. Automatic queue downgrade protection remains unchanged. Full automatic
platform updating and production beta-channel policy remain future work.

Final package validation: 19 isolated installer lifecycle checks passed, including
installation of the actual previous 0.7.0 package, replacement by 0.7.1, changed
loader hash, preservation of the previous payload, reinstall and uninstall.
Signed delivery: artifacts/loader-installer/0.7.1. No live ribbon test claimed.
