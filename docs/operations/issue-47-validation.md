# Dimension length and Exchange placement acceptance

Source change, 2026-11-11. Issue:
https://github.com/statisticsnorway/nadabas/issues/47

Create Key Family previously accepted any numeric text in Add. Values above
255, fractional values and Long overflows could reach conversion or schema work.
`DbCreateTables.TryDimensionLength` now accepts whole numbers from 1 through 255
and rejects invalid input before conversion to Long. Add and Finish use it,
with a message in English, French, Portuguese and Bahasa Indonesia.
`CreateNewKeyFam` validates all dimension lengths before preparing any table.
The existing ordering creates the table before calling `CurrentDB.AddKeyname`.

This prevents invalid dimension lengths from starting creation. It does not
make all schema operations transactional or repair earlier partial creations.
Do not remove existing metadata automatically: inspect the affected database
and agree recovery from a verified backup separately.

Both ribbon schemas and the beta patch tool place Show Exchange tools directly
after Create Exchange links in Administration / Other admin. The administrator
group and existing active-workbook link condition govern availability. The
preference still controls display only and does not disable Exchange connections.

## Evidence and remaining acceptance

- The exact production parser ran in a disposable Excel workbook with 14 cases,
  covering 1, 255, 256, fractions, empty text, invalid text and large values.
- Source regression checks verify Add, Finish and the schema-entry guard ordering.
- The VBA importer dry-run validated 191 source components against the template.
- No database was opened or changed by the runtime parser test.

Before signing a new VBA payload, import the reviewed source and updated language
resources/ribbon into a test copy. In a disposable Access database, try 256 and
invalid inputs in Create Key Family, including an invalid second dimension.
Verify that neither a table nor a KeyNames entry is created. Correct the input
and reuse the same name successfully. Test boundary lengths 1 and 255, within
the backend's total index-width constraints. Repeat creation on the supported
SQL Server configuration before claiming that backend acceptance.

Inspect the ribbon as an administrator with and without Exchange links. Confirm
the toggle is adjacent to Create Exchange links, preserves its preference, and
shows/hides the existing Exchange groups. Confirm the administrator-only parent
visibility. Check all four languages and keyboard operation.

These source changes are not present in previously signed packages. VBA changes
require a newly built and signed payload; updating the .NET loader alone does
not include this fix.
