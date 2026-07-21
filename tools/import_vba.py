"""Import version-controlled VBA source files into a copy of a NADABAS XLAM.

The importer uses Excel's VBA extensibility API.  Existing components are
updated by replacing their code module, which deliberately preserves the
designer and binary state of UserForms in the template XLAM.  A new UserForm
can only be imported from a complete ``.frm`` export (and every referenced
``.frx`` sidecar must be present).

Excel must be installed and "Trust access to the VBA project object model"
must be enabled for the user running this command.
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
from typing import Any, Iterable, Sequence


VBEXT_CT_STANDARD_MODULE = 1
VBEXT_CT_CLASS_MODULE = 2
VBEXT_CT_MS_FORM = 3
VBEXT_CT_DOCUMENT = 100
VBEXT_PP_NONE = 0
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

VBA_SUFFIXES = {".bas", ".cls", ".frm"}
KIND_ORDER = {"module": 0, "class": 1, "form": 2}
ATTRIBUTE_NAME_RE = re.compile(
    r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$', re.IGNORECASE | re.MULTILINE
)
FRX_REFERENCE_RE = re.compile(r'"([^"]+\.frx)"\s*:', re.IGNORECASE)


class VbaImportError(RuntimeError):
    """Raised when the VBA sources cannot be imported safely."""


@dataclass(frozen=True)
class SourceComponent:
    """A validated VBA source file."""

    path: Path
    name: str
    kind: str
    code: str
    complete_export: bool
    frx_paths: tuple[Path, ...] = ()


@dataclass(frozen=True)
class ExistingComponent:
    """The small part of a VBComponent needed to plan an import."""

    name: str
    component_type: int


@dataclass(frozen=True)
class ImportAction:
    """One planned change to the VBA project."""

    operation: str
    source: SourceComponent


@dataclass(frozen=True)
class ImportReport:
    """Result returned by :func:`import_vba`."""

    template: Path
    output: Path | None
    actions: tuple[ImportAction, ...]
    dry_run: bool


def _normalise_newlines(text: str) -> str:
    # Some historical exports contain CRCRLF.  Treat it as one line ending.
    return text.replace("\r\r\n", "\n").replace("\r\n", "\n").replace("\r", "\n")


def _read_vba_text(path: Path) -> str:
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return _normalise_newlines(data.decode(encoding))
        except UnicodeDecodeError:
            continue
    raise VbaImportError(f"Could not decode VBA source: {path}")


def _component_kind(path: Path) -> str:
    return {".bas": "module", ".cls": "class", ".frm": "form"}[path.suffix.lower()]


def _component_name(path: Path, text: str) -> str:
    match = ATTRIBUTE_NAME_RE.search(text)
    name = match.group(1) if match else path.stem
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name):
        raise VbaImportError(f"Invalid VBA component name {name!r} in {path}")
    return name


def _is_complete_export(kind: str, text: str) -> bool:
    first = next((line.strip() for line in text.splitlines() if line.strip()), "")
    if kind == "form":
        return first.upper().startswith("VERSION ") and bool(
            re.search(r"^\s*Begin\s+VB\.UserForm\b", text, re.IGNORECASE | re.MULTILINE)
        )
    if kind == "class":
        return first.upper().startswith("VERSION ")
    # Exported standard modules legitimately start with Attribute VB_Name.
    return bool(ATTRIBUTE_NAME_RE.search(text))


def _extract_code(path: Path, kind: str, text: str) -> str:
    """Return code accepted by CodeModule.AddFromString.

    Export-only metadata and a possible UserForm designer block are excluded.
    In a regular exported component, executable code follows the final
    ``Attribute`` line.
    """

    lines = text.splitlines()
    attribute_indexes = [
        index for index, line in enumerate(lines) if line.lstrip().lower().startswith("attribute ")
    ]
    if attribute_indexes:
        code_lines = lines[max(attribute_indexes) + 1 :]
    elif kind == "form" and _is_complete_export(kind, text):
        raise VbaImportError(
            f"Complete UserForm export has no Attribute block; cannot isolate code safely: {path}"
        )
    else:
        code_lines = [line for line in lines if not line.strip().upper().startswith("VERSION ")]

    while code_lines and not code_lines[0].strip():
        code_lines.pop(0)
    while code_lines and not code_lines[-1].strip():
        code_lines.pop()
    return "\n".join(code_lines) + ("\n" if code_lines else "")


def _referenced_frx_files(path: Path, text: str) -> tuple[Path, ...]:
    result: list[Path] = []
    for raw_reference in FRX_REFERENCE_RE.findall(text):
        reference = Path(raw_reference.replace("\\", "/")).name
        candidate = path.with_name(reference)
        if candidate not in result:
            result.append(candidate)
    return tuple(result)


def parse_source(path: Path) -> SourceComponent:
    """Parse and validate one ``.bas``, ``.cls`` or ``.frm`` file."""

    path = path.resolve()
    if path.suffix.lower() not in VBA_SUFFIXES:
        raise VbaImportError(f"Unsupported VBA source extension: {path}")

    text = _read_vba_text(path)
    kind = _component_kind(path)
    frx_paths = _referenced_frx_files(path, text)
    missing = [sidecar for sidecar in frx_paths if not sidecar.is_file()]
    if missing:
        names = ", ".join(str(sidecar) for sidecar in missing)
        raise VbaImportError(f"Missing UserForm binary sidecar(s) referenced by {path}: {names}")

    return SourceComponent(
        path=path,
        name=_component_name(path, text),
        kind=kind,
        code=_extract_code(path, kind, text),
        complete_export=_is_complete_export(kind, text),
        frx_paths=frx_paths,
    )


def discover_sources(source_root: str | Path) -> tuple[SourceComponent, ...]:
    """Discover VBA components below ``source_root`` and reject duplicate names."""

    root = Path(source_root).resolve()
    if not root.is_dir():
        raise VbaImportError(f"VBA source directory does not exist: {root}")

    sources = [
        parse_source(path)
        for path in sorted(root.rglob("*"), key=lambda item: str(item).casefold())
        if path.is_file() and path.suffix.lower() in VBA_SUFFIXES
    ]
    if not sources:
        raise VbaImportError(f"No VBA source files found below: {root}")

    seen: dict[str, Path] = {}
    for source in sources:
        key = source.name.casefold()
        if key in seen:
            raise VbaImportError(
                f"Duplicate VBA component name {source.name!r}: {seen[key]} and {source.path}"
            )
        seen[key] = source.path

    return tuple(sorted(sources, key=lambda item: (KIND_ORDER[item.kind], item.name.casefold())))


def _expected_type(source: SourceComponent) -> int:
    return {
        "module": VBEXT_CT_STANDARD_MODULE,
        "class": VBEXT_CT_CLASS_MODULE,
        "form": VBEXT_CT_MS_FORM,
    }[source.kind]


def plan_import(
    sources: Sequence[SourceComponent],
    existing: Iterable[ExistingComponent],
    *,
    replace_form_designers: bool = False,
) -> tuple[ImportAction, ...]:
    """Plan safe component updates without touching Excel."""

    existing_by_name = {component.name.casefold(): component for component in existing}
    actions: list[ImportAction] = []

    for source in sources:
        current = existing_by_name.get(source.name.casefold())
        if current:
            allowed_types = {_expected_type(source)}
            if source.kind == "class":
                allowed_types.add(VBEXT_CT_DOCUMENT)
            if current.component_type not in allowed_types:
                raise VbaImportError(
                    f"Component type mismatch for {source.name}: source is {source.kind}, "
                    f"but the XLAM component type is {current.component_type}"
                )
            if (
                source.kind == "form"
                and replace_form_designers
                and source.complete_export
            ):
                operation = "replace-component"
            else:
                operation = "replace-code"
        else:
            if source.kind == "form" and not source.complete_export:
                raise VbaImportError(
                    f"UserForm {source.name} is not present in the template and {source.path} "
                    "is code-only. Export the complete .frm and its .frx file from Excel first."
                )
            if source.kind == "class" and not source.complete_export and re.search(
                r"^\s*Attribute\s+VB_PredeclaredId\s*=\s*True\s*$",
                _read_vba_text(source.path),
                re.IGNORECASE | re.MULTILINE,
            ):
                raise VbaImportError(
                    f"New class {source.name} requires VB_PredeclaredId=True, but {source.path} "
                    "is not a complete class export. Export the complete .cls from Excel first."
                )
            operation = "import-component" if source.complete_export else "create-code-component"
        actions.append(ImportAction(operation=operation, source=source))

    return tuple(actions)


def _load_com_modules() -> tuple[Any, Any]:
    if sys.platform != "win32":
        raise VbaImportError("VBA import requires Windows with Microsoft Excel installed.")
    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise VbaImportError(
            "pywin32 is required. Install the project dependencies and try again."
        ) from exc
    return pythoncom, win32com.client


def _project_components(project: Any) -> tuple[ExistingComponent, ...]:
    return tuple(
        ExistingComponent(
            name=str(project.VBComponents.Item(index).Name),
            component_type=int(project.VBComponents.Item(index).Type),
        )
        for index in range(1, int(project.VBComponents.Count) + 1)
    )


def _component_by_name(project: Any, name: str) -> Any:
    for index in range(1, int(project.VBComponents.Count) + 1):
        component = project.VBComponents.Item(index)
        if str(component.Name).casefold() == name.casefold():
            return component
    raise VbaImportError(f"VBA component disappeared during import: {name}")


def _replace_component_code(component: Any, code: str) -> None:
    code_module = component.CodeModule
    line_count = int(code_module.CountOfLines)
    if line_count:
        code_module.DeleteLines(1, line_count)
    if code:
        code_module.AddFromString(code)


def _apply_actions(project: Any, actions: Sequence[ImportAction]) -> None:
    for action in actions:
        source = action.source
        if action.operation == "replace-code":
            _replace_component_code(_component_by_name(project, source.name), source.code)
        elif action.operation == "replace-component":
            project.VBComponents.Remove(_component_by_name(project, source.name))
            imported = project.VBComponents.Import(str(source.path))
            if str(imported.Name).casefold() != source.name.casefold():
                raise VbaImportError(
                    f"Excel imported {source.path} as {imported.Name!r}, expected {source.name!r}"
                )
        elif action.operation == "import-component":
            imported = project.VBComponents.Import(str(source.path))
            if str(imported.Name).casefold() != source.name.casefold():
                raise VbaImportError(
                    f"Excel imported {source.path} as {imported.Name!r}, expected {source.name!r}"
                )
        elif action.operation == "create-code-component":
            component = project.VBComponents.Add(_expected_type(source))
            component.Name = source.name
            _replace_component_code(component, source.code)
        else:  # pragma: no cover - guarded by plan_import
            raise AssertionError(f"Unknown import action: {action.operation}")


def _open_excel_workbook(excel: Any, path: Path, *, read_only: bool) -> Any:
    return excel.Workbooks.Open(
        str(path),
        UpdateLinks=0,
        ReadOnly=read_only,
        IgnoreReadOnlyRecommended=True,
        AddToMru=False,
    )


def _get_unprotected_project(workbook: Any) -> Any:
    try:
        project = workbook.VBProject
        protection = int(project.Protection)
    except Exception as exc:
        raise VbaImportError(
            "Excel denied access to the VBA project. Enable 'Trust access to the VBA "
            "project object model' in Excel Trust Center, then try again."
        ) from exc
    if protection != VBEXT_PP_NONE:
        raise VbaImportError("The VBA project is locked. Unlock it before importing source files.")
    return project


def _validate_paths(
    template_xlam: str | Path,
    output_xlam: str | Path | None,
    *,
    dry_run: bool,
    force: bool,
) -> tuple[Path, Path | None]:
    template = Path(template_xlam).resolve()
    if not template.is_file():
        raise VbaImportError(f"Template XLAM does not exist: {template}")
    if template.suffix.lower() != ".xlam":
        raise VbaImportError(f"Template must be an .xlam file: {template}")

    if dry_run and output_xlam is None:
        return template, None
    output = (
        Path(output_xlam).resolve()
        if output_xlam is not None
        else template.with_name(f"{template.stem}-rebuilt.xlam")
    )
    if output.suffix.lower() != ".xlam":
        raise VbaImportError(f"Output must be an .xlam file: {output}")
    if output == template:
        raise VbaImportError("Output must be different from the template XLAM.")
    if output.exists() and not force and not dry_run:
        raise VbaImportError(f"Output already exists (use --force to replace it): {output}")
    return template, output


def import_vba(
    template_xlam: str | Path,
    source_root: str | Path = "vba",
    output_xlam: str | Path | None = None,
    *,
    dry_run: bool = False,
    force: bool = False,
    visible: bool = False,
    replace_form_designers: bool = False,
) -> ImportReport:
    """Import VBA sources into a new XLAM and return an action report.

    The template is never modified.  The completed workbook replaces ``output``
    only after Excel has saved it successfully.
    """

    template, output = _validate_paths(
        template_xlam, output_xlam, dry_run=dry_run, force=force
    )
    sources = discover_sources(source_root)
    pythoncom, win32_client = _load_com_modules()

    excel = None
    workbook = None
    staging: Path | None = None
    pythoncom.CoInitialize()
    try:
        excel = win32_client.DispatchEx("Excel.Application")
        excel.Visible = visible
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

        if dry_run:
            workbook = _open_excel_workbook(excel, template, read_only=True)
            project = _get_unprotected_project(workbook)
            actions = plan_import(
                sources,
                _project_components(project),
                replace_form_designers=replace_form_designers,
            )
            return ImportReport(template, output, actions, True)

        assert output is not None
        output.parent.mkdir(parents=True, exist_ok=True)
        staging = output.with_name(f".{output.stem}.{uuid.uuid4().hex}.tmp.xlam")
        shutil.copy2(template, staging)
        workbook = _open_excel_workbook(excel, staging, read_only=False)
        project = _get_unprotected_project(workbook)
        actions = plan_import(
            sources,
            _project_components(project),
            replace_form_designers=replace_form_designers,
        )
        _apply_actions(project, actions)
        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None
        os.replace(staging, output)
        staging = None
        return ImportReport(template, output, actions, False)
    except VbaImportError:
        raise
    except Exception as exc:
        raise VbaImportError(f"Excel failed while importing VBA: {exc}") from exc
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
        if staging is not None:
            try:
                staging.unlink(missing_ok=True)
            except OSError:
                pass
        pythoncom.CoUninitialize()


def _print_report(report: ImportReport) -> None:
    label = "DRY RUN" if report.dry_run else "BUILT"
    print(f"{label}: {report.template}")
    if report.output:
        print(f"Output: {report.output}")
    print(f"Components: {len(report.actions)}")
    for action in report.actions:
        print(f"  {action.operation:22} {action.source.name} ({action.source.path.name})")
    if not report.dry_run:
        print("Warning: changing VBA invalidates any existing digital signature.")


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Import version-controlled NADABAS VBA sources into a copy of an XLAM."
    )
    parser.add_argument("template", type=Path, help="Existing NADABAS .xlam used as template")
    parser.add_argument(
        "--source-root", type=Path, default=Path("vba"), help="VBA source root (default: vba)"
    )
    parser.add_argument("--output", type=Path, help="Output .xlam path")
    parser.add_argument("--dry-run", action="store_true", help="Validate and show changes only")
    parser.add_argument("--force", action="store_true", help="Replace an existing output file")
    parser.add_argument("--visible", action="store_true", help="Show Excel while importing")
    parser.add_argument(
        "--replace-form-designers",
        action="store_true",
        help="Replace existing form designers from complete .frm/.frx exports",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        report = import_vba(
            args.template,
            args.source_root,
            args.output,
            dry_run=args.dry_run,
            force=args.force,
            visible=args.visible,
            replace_form_designers=args.replace_form_designers,
        )
    except VbaImportError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    _print_report(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
