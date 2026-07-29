"""Export a complete VBA project from a NADABAS Excel file.

Excel must be installed and "Trust access to the VBA project object model"
must be enabled for the user running this command.  The source workbook is
opened read-only with macros disabled.  Components are exported to a staging
directory so a failed export never leaves a partial output tree behind.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from typing import Sequence

VBEXT_CT_STANDARD_MODULE = 1
VBEXT_CT_CLASS_MODULE = 2
VBEXT_CT_MS_FORM = 3
VBEXT_CT_DOCUMENT = 100
VBEXT_PP_NONE = 0
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

SUPPORTED_WORKBOOK_SUFFIXES = {".xlam", ".xlsm", ".xlsb"}
COMPONENT_LAYOUT = {
    VBEXT_CT_STANDARD_MODULE: ("modules", ".bas"),
    VBEXT_CT_CLASS_MODULE: ("classes", ".cls"),
    VBEXT_CT_MS_FORM: ("forms", ".frm"),
    VBEXT_CT_DOCUMENT: ("classes", ".cls"),
}
COMPONENT_NAME_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]*\Z")


class VbaExportError(RuntimeError):
    """Raised when a complete VBA export cannot be produced safely."""


@dataclass(frozen=True)
class ExportedComponent:
    """One component exported from the VBA project."""

    name: str
    component_type: int
    path: Path
    sidecars: tuple[Path, ...] = ()


@dataclass(frozen=True)
class ExportReport:
    """Result returned by :func:`export_vba`."""

    workbook: Path
    output_root: Path
    components: tuple[ExportedComponent, ...]


def component_relative_path(name: str, component_type: int) -> Path:
    """Return the repository path for one supported VBA component."""

    if not COMPONENT_NAME_RE.fullmatch(name):
        raise VbaExportError(f"Unsafe VBA component name: {name!r}")
    try:
        directory, suffix = COMPONENT_LAYOUT[component_type]
    except KeyError as exc:
        raise VbaExportError(
            f"Unsupported VBA component type {component_type} for {name!r}"
        ) from exc
    return Path(directory) / f"{name}{suffix}"


def _load_com_modules() -> tuple[Any, Any]:
    if sys.platform != "win32":
        raise VbaExportError(
            "VBA export requires Windows with Microsoft Excel installed."
        )
    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise VbaExportError(
            "pywin32 is required. Install the project dependencies and try again."
        ) from exc
    return pythoncom, win32com.client


def _validate_paths(
    workbook_path: str | Path,
    output_root: str | Path,
    *,
    force: bool,
) -> tuple[Path, Path]:
    workbook = Path(workbook_path).resolve()
    output = Path(output_root).resolve()

    if not workbook.is_file():
        raise VbaExportError(f"Workbook does not exist: {workbook}")
    if workbook.suffix.lower() not in SUPPORTED_WORKBOOK_SUFFIXES:
        allowed = ", ".join(sorted(SUPPORTED_WORKBOOK_SUFFIXES))
        raise VbaExportError(f"Workbook must be one of {allowed}: {workbook}")
    if output.exists() and not force:
        raise VbaExportError(
            f"Output already exists (use --force to replace it): {output}"
        )
    if output == workbook or workbook in output.parents:
        raise VbaExportError("Output directory must not contain the source workbook.")
    return workbook, output


def _get_unprotected_project(workbook: Any) -> Any:
    try:
        project = workbook.VBProject
        protection = int(project.Protection)
    except Exception as exc:
        raise VbaExportError(
            "Excel denied access to the VBA project. Enable 'Trust access to the VBA "
            "project object model' in Excel Trust Center, then try again."
        ) from exc
    if protection != VBEXT_PP_NONE:
        raise VbaExportError(
            "The VBA project is locked. Unlock it before exporting source."
        )
    return project


def _project_components(project: Any) -> list[Any]:
    components = [
        project.VBComponents.Item(index)
        for index in range(1, int(project.VBComponents.Count) + 1)
    ]
    return sorted(
        components,
        key=lambda component: (int(component.Type), str(component.Name).casefold()),
    )


def _plan_destinations(components: Sequence[Any]) -> list[tuple[Any, Path]]:
    planned: list[tuple[Any, Path]] = []
    seen: dict[str, str] = {}

    for component in components:
        name = str(component.Name)
        relative_path = component_relative_path(name, int(component.Type))
        key = str(relative_path).casefold()
        if key in seen:
            raise VbaExportError(
                f"Duplicate export destination for {name!r} and {seen[key]!r}: "
                f"{relative_path}"
            )
        seen[key] = name
        planned.append((component, relative_path))
    return planned


def _normalise_text_export(path: Path) -> None:
    """Match repository whitespace rules without changing VBA semantics."""

    data = path.read_bytes()
    encoding = "utf-8-sig" if data.startswith(b"\xef\xbb\xbf") else "cp1252"
    try:
        text = data.decode(encoding)
    except UnicodeDecodeError as exc:
        raise VbaExportError(f"Could not decode Excel export {path}: {exc}") from exc

    lines = [
        line.rstrip(" \t")
        for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    ]
    while lines and not lines[-1]:
        lines.pop()
    # The VBComponents importer treats LF-only UserForm exports as ordinary
    # standard modules.  Keep all exported VBA text on Windows CRLF endings.
    normalised = "\r\n".join(lines) + "\r\n"
    path.write_bytes(normalised.encode(encoding))


def _export_components(project: Any, staging: Path) -> tuple[ExportedComponent, ...]:
    planned = _plan_destinations(_project_components(project))
    if not planned:
        raise VbaExportError("The VBA project contains no exportable components.")

    for directory in {path.parent for _, path in planned}:
        (staging / directory).mkdir(parents=True, exist_ok=True)

    exported: list[ExportedComponent] = []
    for component, relative_path in planned:
        destination = staging / relative_path
        try:
            component.Export(str(destination))
        except Exception as exc:
            raise VbaExportError(
                f"Excel failed to export component {component.Name!r} to {destination}: {exc}"
            ) from exc
        if not destination.is_file():
            raise VbaExportError(
                f"Excel reported success but did not create the export: {destination}"
            )
        _normalise_text_export(destination)

        sidecars: tuple[Path, ...] = ()
        if destination.suffix.lower() == ".frm":
            sidecars = tuple(sorted(destination.parent.glob(f"{destination.stem}.*")))
            sidecars = tuple(path for path in sidecars if path != destination)

        exported.append(
            ExportedComponent(
                name=str(component.Name),
                component_type=int(component.Type),
                path=relative_path,
                sidecars=tuple(path.relative_to(staging) for path in sidecars),
            )
        )
    return tuple(exported)


def export_vba(
    workbook_path: str | Path,
    output_root: str | Path,
    *,
    force: bool = False,
    visible: bool = False,
) -> ExportReport:
    """Export all supported VBA components from a workbook."""

    workbook_path, output_root = _validate_paths(
        workbook_path, output_root, force=force
    )
    pythoncom, win32_client = _load_com_modules()

    output_root.parent.mkdir(parents=True, exist_ok=True)
    staging = output_root.with_name(f".{output_root.name}.{uuid.uuid4().hex}.tmp")
    excel = None
    workbook = None
    pythoncom.CoInitialize()
    try:
        staging.mkdir(parents=False)
        excel = win32_client.DispatchEx("Excel.Application")
        excel.Visible = visible
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

        workbook = excel.Workbooks.Open(
            str(workbook_path),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )
        project = _get_unprotected_project(workbook)
        components = _export_components(project, staging)

        workbook.Close(SaveChanges=False)
        workbook = None
        if output_root.exists():
            shutil.rmtree(output_root)
        os.replace(staging, output_root)
        return ExportReport(workbook_path, output_root, components)
    except VbaExportError:
        raise
    except Exception as exc:
        raise VbaExportError(f"Excel failed while exporting VBA: {exc}") from exc
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
        if staging.exists():
            try:
                shutil.rmtree(staging)
            except OSError:
                pass
        pythoncom.CoUninitialize()


def _print_report(report: ExportReport) -> None:
    print(f"EXPORTED: {report.workbook}")
    print(f"Output: {report.output_root}")
    print(f"Components: {len(report.components)}")
    for component in report.components:
        suffix = ""
        if component.sidecars:
            suffix = " + " + ", ".join(str(path) for path in component.sidecars)
        print(f"  {component.path}{suffix}")


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Export a complete NADABAS VBA project through Microsoft Excel."
    )
    parser.add_argument(
        "workbook", type=Path, help="Source .xlam, .xlsm, or .xlsb file"
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=Path("exported-vba"),
        help="New output directory (default: exported-vba)",
    )
    parser.add_argument(
        "--force", action="store_true", help="Replace an existing output"
    )
    parser.add_argument(
        "--visible", action="store_true", help="Show Excel during export"
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        report = export_vba(
            args.workbook,
            args.output_root,
            force=args.force,
            visible=args.visible,
        )
    except VbaExportError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    _print_report(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
