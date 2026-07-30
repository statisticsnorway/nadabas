# Scan exported VBA for security-relevant operations

`tools/scan_vba_security.py` is a static review aid for the exported `.bas`,
`.cls`, and `.frm` files below `vba/`. It reports code that deserves human
review; a finding does not by itself mean that code is malicious.

## Run the scanner

From the repository root:

```bash
python tools/scan_vba_security.py vba
```

The default policy returns exit code `1` only for unapproved `error` findings.
Warnings and information remain visible without failing CI. Use a stricter
threshold when needed:

```bash
python tools/scan_vba_security.py vba --fail-on warning
```

Available exit codes are:

- `0`: scan completed and the selected threshold passed
- `1`: at least one unapproved finding met the selected threshold
- `2`: invalid arguments, input, or approval configuration

Use JSON for CI artifacts or further processing:

```bash
python tools/scan_vba_security.py vba --format json > scan-vba-security.json
```

## Severity model

- `error`: command execution, process termination, downloads, or deletion
- `warning`: shell opening, WMI/process creation, file writes/copies,
  automation hooks, synthetic keys, or network/file-system automation objects
- `info`: native API declarations, environment reads, WMI queries, or directory
  creation

The scanner understands VBA apostrophe and `Rem` comments. It also masks
ordinary string contents for rules that look for VBA identifiers, which avoids
matches caused by comments and UserForm captions. Rules that must inspect
object names or commands inside strings, such as `CreateObject` and command
interpreters, still inspect them deliberately.

## Document a reviewed exception

Do not suppress an entire VBA file. If a finding is intentional and has been
reviewed, create a JSON file that matches the exact path, rule, and expected
line shape:

```json
{
  "approvals": [
    {
      "path": "vba/modules/InstallMeAsAddin.bas",
      "rule_id": "file-delete",
      "line_regex": "^\\s*Kill\\s+AI\\.FullName$",
      "reason": "Reviewed installer operation that replaces the selected NADABAS add-in."
    }
  ]
}
```

Then run:

```bash
python tools/scan_vba_security.py vba \
  --approvals tools/vba-security-approvals.json \
  --show-approved
```

Keep the approval file under version control. The line expression should be as
narrow as practical so a future change of target or operation becomes a new,
unapproved finding.

The repository approval file is `tools/vba-security-approvals.json`. GitHub
Actions loads it explicitly and prints approved operations alongside warnings
and informational findings. An approval documents reviewed, necessary
behaviour; it does not disable the rule for other lines or files.

## Run the tests

```bash
python -m unittest discover -s tests -p "test_scan_vba_security.py" -v
```

Run all repository tests with the project source directory on the Python path:

```bash
PYTHONPATH=src python -m unittest discover -s tests -v
```
