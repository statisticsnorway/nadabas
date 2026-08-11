"""Apply reviewed Ribbon files and language rows to a copy of a NADABAS XLAM.

The input add-in is never modified. Excel is used only to update the hidden
language worksheets so VBA, form designers, and workbook metadata remain
intact. The reviewed ``customUI`` package files are then copied byte-for-byte
into the resulting Office Open XML package.
"""

from __future__ import annotations

import argparse
import csv
import os
import shutil
import sys
import uuid
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from typing import Sequence


MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
RESOURCE_FILE = Path("resources/office-ui.csv")
CUSTOM_UI_DIRECTORY = Path("customUI")


class OfficeUiError(RuntimeError):
    """Raised when UI resources cannot be applied safely."""


@dataclass(frozen=True)
class ResourceRow:
    sheet: str
    key: str
    control: str
    english: str
    french: str
    portuguese: str


def read_resource_rows(path: str | Path) -> tuple[ResourceRow, ...]:
    source = Path(path)
    if not source.is_file():
        raise OfficeUiError(f"Language resource file does not exist: {source}")

    with source.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        expected = {
            "sheet",
            "key",
            "control",
            "english",
            "french",
            "portuguese",
        }
        if set(reader.fieldnames or ()) != expected:
            raise OfficeUiError(
                f"Unexpected columns in {source}; expected {sorted(expected)}"
            )
        rows = tuple(ResourceRow(**row) for row in reader)

    if not rows:
        raise OfficeUiError(f"Language resource file is empty: {source}")

    seen: set[tuple[str, str, str]] = set()
    for row in rows:
        if row.sheet not in {"Ribbon", "Forms", "Messages"}:
            raise OfficeUiError(f"Unsupported language sheet: {row.sheet!r}")
        if not row.key or (row.sheet == "Forms" and not row.control):
            raise OfficeUiError(f"Incomplete language resource row: {row}")
        identity = (row.sheet.casefold(), row.key.casefold(), row.control.casefold())
        if identity in seen:
            raise OfficeUiError(f"Duplicate language resource row: {row}")
        seen.add(identity)
    return rows


def _load_com_modules() -> tuple[Any, Any]:
    if sys.platform != "win32":
        raise OfficeUiError("Applying XLAM language resources requires Windows.")
    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise OfficeUiError(
            "pywin32 is required. Install the project dependencies and try again."
        ) from exc
    return pythoncom, win32com.client


def _resource_row_number(sheet: Any, resource: ResourceRow) -> int:
    first_row = int(sheet.UsedRange.Row)
    last_row = first_row + int(sheet.UsedRange.Rows.Count) - 1

    for row_number in range(first_row, last_row + 1):
        key = str(sheet.Cells(row_number, 1).Value or "").strip()
        control = str(sheet.Cells(row_number, 2).Value or "").strip()
        if key.casefold() != resource.key.casefold():
            continue
        if (
            resource.sheet == "Forms"
            and control.casefold() != resource.control.casefold()
        ):
            continue
        return row_number
    return last_row + 1


def _apply_language_resources(workbook_path: Path, rows: Sequence[ResourceRow]) -> None:
    pythoncom, win32_client = _load_com_modules()
    excel = None
    workbook = None

    pythoncom.CoInitialize()
    try:
        excel = win32_client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        workbook = excel.Workbooks.Open(
            str(workbook_path),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )

        for resource in rows:
            try:
                sheet = workbook.Worksheets(resource.sheet)
            except Exception as exc:
                raise OfficeUiError(
                    f"Workbook has no {resource.sheet!r} language sheet."
                ) from exc

            row_number = _resource_row_number(sheet, resource)
            sheet.Cells(row_number, 1).Value = resource.key
            sheet.Cells(row_number, 2).Value = resource.control
            sheet.Cells(row_number, 3).Value = resource.english
            sheet.Cells(row_number, 4).Value = resource.french
            sheet.Cells(row_number, 5).Value = resource.portuguese

        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None
    except OfficeUiError:
        raise
    except Exception as exc:
        raise OfficeUiError(
            f"Excel failed while applying language resources: {exc}"
        ) from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def custom_ui_replacements(source_directory: str | Path) -> dict[str, bytes]:
    source = Path(source_directory)
    if not source.is_dir():
        raise OfficeUiError(f"customUI source directory does not exist: {source}")

    replacements = {
        path.relative_to(source.parent).as_posix(): path.read_bytes()
        for path in sorted(source.rglob("*"))
        if path.is_file()
    }
    required = {"customUI/customUI.xml", "customUI/customUI14.xml"}
    missing = required.difference(replacements)
    if missing:
        raise OfficeUiError(f"Missing required Ribbon files: {sorted(missing)}")
    return replacements


def replace_custom_ui(package_path: str | Path, source_directory: str | Path) -> None:
    package = Path(package_path)
    replacements = custom_ui_replacements(source_directory)
    temporary = package.with_name(f".{package.name}.{uuid.uuid4().hex}.tmp")

    try:
        with zipfile.ZipFile(package, "r") as source_zip:
            package_names = set(source_zip.namelist())
            absent = set(replacements).difference(package_names)
            if absent:
                raise OfficeUiError(
                    f"XLAM package is missing reviewed customUI parts: {sorted(absent)}"
                )

            with zipfile.ZipFile(temporary, "w", allowZip64=True) as target_zip:
                for info in source_zip.infolist():
                    data = replacements.get(
                        info.filename, source_zip.read(info.filename)
                    )
                    target_zip.writestr(info, data)
        os.replace(temporary, package)
    finally:
        temporary.unlink(missing_ok=True)


def apply_office_ui(
    template_xlam: str | Path,
    output_xlam: str | Path,
    source_root: str | Path = "vba",
    *,
    force: bool = False,
) -> Path:
    template = Path(template_xlam).resolve()
    output = Path(output_xlam).resolve()
    source = Path(source_root).resolve()

    if not template.is_file() or template.suffix.lower() != ".xlam":
        raise OfficeUiError(f"Input must be an existing .xlam file: {template}")
    if output.suffix.lower() != ".xlam" or output == template:
        raise OfficeUiError("Output must be a different .xlam path.")
    if output.exists() and not force:
        raise OfficeUiError(f"Output already exists (use --force): {output}")

    rows = read_resource_rows(source / RESOURCE_FILE)
    custom_ui_replacements(source / CUSTOM_UI_DIRECTORY)

    output.parent.mkdir(parents=True, exist_ok=True)
    staging = output.with_name(f".{output.stem}.{uuid.uuid4().hex}.tmp.xlam")
    try:
        shutil.copy2(template, staging)
        _apply_language_resources(staging, rows)
        replace_custom_ui(staging, source / CUSTOM_UI_DIRECTORY)
        os.replace(staging, output)
    finally:
        staging.unlink(missing_ok=True)
    return output


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Apply reviewed Ribbon and language resources to an XLAM copy."
    )
    parser.add_argument("template", type=Path, help="Input .xlam file")
    parser.add_argument("--output", type=Path, required=True, help="Output .xlam file")
    parser.add_argument(
        "--source-root", type=Path, default=Path("vba"), help="UI source root"
    )
    parser.add_argument("--force", action="store_true", help="Replace output")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        output = apply_office_ui(
            args.template,
            args.output,
            args.source_root,
            force=args.force,
        )
    except OfficeUiError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print(f"BUILT: {output}")
    print("Warning: changing workbook UI invalidates any existing digital signature.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
