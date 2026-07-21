from pathlib import Path
import unittest

from tools.import_vba import (
    ExistingComponent,
    VbaImportError,
    discover_sources,
    parse_source,
    plan_import,
)


FIXTURES = Path(__file__).with_name("fixtures")


class ImportVbaTests(unittest.TestCase):
    def test_parse_standard_module_removes_export_attributes(self) -> None:
        source = parse_source(FIXTURES / "module" / "Example.bas")

        self.assertEqual(source.name, "Example")
        self.assertEqual(source.kind, "module")
        self.assertTrue(source.complete_export)
        self.assertEqual(
            source.code, 'Option Explicit\nPublic Sub Run()\n    MsgBox "Hei"\nEnd Sub\n'
        )

    def test_code_only_existing_form_preserves_designer(self) -> None:
        source = parse_source(FIXTURES / "code_only_form" / "Dialog.frm")

        actions = plan_import([source], [ExistingComponent("Dialog", 3)])

        self.assertEqual(actions[0].operation, "replace-code")
        self.assertFalse(source.complete_export)

    def test_code_only_new_form_is_rejected(self) -> None:
        source = parse_source(FIXTURES / "code_only_form" / "Dialog.frm")

        with self.assertRaisesRegex(VbaImportError, "code-only"):
            plan_import([source], [])

    def test_complete_form_requires_referenced_frx(self) -> None:
        with self.assertRaisesRegex(VbaImportError, "Missing UserForm binary"):
            parse_source(FIXTURES / "missing_frx" / "Dialog.frm")

    def test_complete_form_can_replace_designer_when_requested(self) -> None:
        path = FIXTURES / "complete_form" / "Dialog.frm"
        source = parse_source(path)

        actions = plan_import(
            [source], [ExistingComponent("Dialog", 3)], replace_form_designers=True
        )

        self.assertEqual(actions[0].operation, "replace-component")
        self.assertEqual(source.frx_paths, (path.with_suffix(".frx").resolve(),))

    def test_duplicate_component_names_are_rejected(self) -> None:
        with self.assertRaisesRegex(VbaImportError, "Duplicate VBA component"):
            discover_sources(FIXTURES / "duplicates")

    def test_document_module_is_updated_as_code(self) -> None:
        source = parse_source(FIXTURES / "document" / "ThisWorkbook.cls")

        actions = plan_import([source], [ExistingComponent("ThisWorkbook", 100)])

        self.assertEqual(actions[0].operation, "replace-code")

    def test_component_type_mismatch_is_rejected(self) -> None:
        source = parse_source(FIXTURES / "module" / "Example.bas")

        with self.assertRaisesRegex(VbaImportError, "type mismatch"):
            plan_import([source], [ExistingComponent("Example", 3)])


if __name__ == "__main__":
    unittest.main()
