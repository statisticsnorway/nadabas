# Access workbook migration compatibility repair

Updated: 2026-11-24.
Status: reproduced on an affected database; migration repair validated on isolated copies. Distribution and ordinary workbook load/save acceptance remain separate.

## Maintainer decision

Do not distribute a diagnostic build. Preserve existing Access databases and the WorkbookID contract. Do not bypass migration or replace approved signed artifacts under their existing identities. Confidential input and extracted database copies stay outside the repository and cloud-synced folders; reports must not contain workbook names or row values.

## Confirmed regression

The supplied 5.10.004 add-in does not call EnsureWorkbookIdentitySchema. The current non-exchange database opening flow runs this migration before loading administrators and domain collections. The affected report identifies 6.02.005 as failing and 5.10.004 as working, with no configured administrators. Changing the mapped drive to a native path did not resolve the reported failure.

The actual affected database now reproduces provider error -2147467259 at `Assign workbook identity: DescriptionDimensions.WorkbookID`. Reading names and resolving identities succeeds; updating the matching legacy rows through the existing name index fails. Reopening the ACE connection does not resolve it. On separate disposable copies, either expressing the name predicate as `([WorkbookName] & '') = ...` or dropping the legacy PrimaryIndex makes that same update succeed. This isolates the failing indexed lookup path; it does not establish the underlying physical cause or identify a deleted business record. Administrator permissions are not the demonstrated cause.

The original index also supplies a COUNT(*) different from the number of rows that can be enumerated. Preservation checks therefore compare complete enumerated row multisets, not just index-based counts. No Compact/Repair is applied.

## Repair

For Access only, backfill compares the legacy name with an empty-string concatenation, avoiding the failing direct name-index lookup. Only NULL identity fields are updated. Existing name escaping, ID generation and the subsequent index rebuild remain in place. SQL Server retains its existing predicate; exchange databases still skip this migration.

The error handler reports the schema step instead of assuming a permission problem or claiming that no partial schema changes occurred. A failed migration-history read is propagated rather than interpreted as an unapplied migration. A failed upgrade does not record the completion marker.

## Evidence

`tools/test_workbook_identity_access.py` executes production migration procedures in an owned Excel process with an unsaved scratch workbook. It uses existing VBA-project access without changing Office policy, user registration or the original files.

- The affected database fails before the backfill predicate change.
- Fresh migration and recovery from the actual partially migrated failure both pass after the change.
- All original non-identity field values and complete row multiplicities across 24 local tables are preserved. Generated identity fields are validated separately against their workbook names.
- Repeated migration succeeds without changing legacy data.
- Synthetic reference tables, explicit injected-failure recovery, database switching and the actual 5.10.004 table-creation fixture pass. The injected failure tests recovery, not reproduction of the affected database defect.
- SQL Server and x86 Office have not been exercised against this confidential database. This repair introduces no architecture-dependent API.

Sanitized local evidence is held in ignored `artifacts/access-migration/confirmed-repair-report.json`. The owned confidential extraction and test copies were removed after validation. The supplied archive was verified unchanged and retained.

## Release acceptance

The original signed 6.02.009 artifact remains unchanged. A separate unsigned review copy imports only WorkbookIdentitySchema and is not a distributable release. Assign the approved release identity, sign final bytes and rebuild the installer through the established release process. Do not distribute the earlier diagnostic artifact.

Migration and data-preservation evidence does not replace a normal Excel open/load/save check with the final signed package. Do not claim full workshop or production acceptance from the isolated migration checks.

## Final local checks

- Python suite: 140 passed, 1 skipped.
- Repository pre-commit checks passed; the new untracked test harness was also checked explicitly.
- VBA security scan: zero errors, with the existing 14 warnings retained.
- Extra synthetic cases with apostrophes, accented names, trailing spaces and unlinked documents passed.
- Unsigned review build read-back matches WorkbookIdentitySchema except for the VBA editor's automatic `.Value` property casing. The signed template hash is unchanged.
- Confidential temporary copies were removed after source/archive integrity checks. Sanitized cleanup evidence is in ignored `artifacts/access-migration/confidential-cleanup-report.json`.

## Signed candidate handoff — 2026-11-24

VBA 6.02.010 has been signed with the approved SSB publisher and timestamp.
Excel confirms the VBA signature, version, fixed build date and repair module.
The original 6.02.009 remains unchanged. Its successor has a distinct compiled
slot/hash recorded in `dotnet/releases/access-repair-6.02.010.json`.

Both private-runtime architectures and startup components are signed; startup
pins the final signed loader bytes. A signed internal 1.0.0 installer containing
6.02.010 was built and verified. Forty-two isolated installer checks passed,
including the actual x86/x64 helpers, language handling, legacy restoration and
selection of the previous/new VBA versions. The .NET suite passed 210 tests;
Python passed 140 with one skip. These are not ordinary Excel acceptance.

The attempted ordinary-profile automated check stopped at its child PowerShell
signature precheck before Setup ran. The previously installed 0.9.5 registration
remains unchanged. Its cause was not diagnosed; do not equate this failed helper
precheck with a successful desktop test or with an invalid package signature.
Direct build-time publisher verification succeeded.

The maintainer then chose to perform installation and Excel tests locally.
No further desktop tests were started. Keep that workflow in future handoffs.
Signed artifacts, installer identity and local reports are under ignored
`artifacts/access-migration/release-6.02.010/`. Publication remains unapproved.
