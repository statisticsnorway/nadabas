import unittest
from pathlib import Path
from types import SimpleNamespace

from tools.import_vba import ExistingComponent
from tools.import_vba import VbaImportError
from tools.import_vba import _is_complete_export
from tools.import_vba import _repair_and_validate_references
from tools.import_vba import _replace_component_code
from tools.import_vba import discover_sources
from tools.import_vba import parse_source
from tools.import_vba import plan_import

FIXTURES = Path(__file__).with_name("fixtures")
REPOSITORY_ROOT = Path(__file__).resolve().parents[1]


class ImportVbaTests(unittest.TestCase):
    def test_parse_standard_module_removes_export_attributes(self) -> None:
        source = parse_source(FIXTURES / "module" / "Example.bas")

        self.assertEqual(source.name, "Example")
        self.assertEqual(source.kind, "module")
        self.assertTrue(source.complete_export)
        self.assertEqual(
            source.code,
            'Option Explicit\nPublic Sub Run()\n    MsgBox "Hei"\nEnd Sub\n',
        )

    def test_member_attribute_does_not_remove_withevents_declaration(self) -> None:
        source = parse_source(REPOSITORY_ROOT / "vba/classes/AppEvents.cls")

        self.assertIn("Public WithEvents App As Application\n", source.code)
        self.assertNotIn("Attribute App.VB_VarHelpID", source.code)

    def test_member_attributes_preserve_all_clsnode_declarations(self) -> None:
        source = parse_source(REPOSITORY_ROOT / "vba/classes/clsNode.cls")

        for declaration in (
            "Private WithEvents mctlControl As MSForms.label",
            "Private WithEvents mctlExpander As MSForms.label",
            "Private WithEvents moEditBox As MSForms.TextBox",
            "Private WithEvents mctlCheckBox As MSForms.label",
        ):
            self.assertIn(declaration, source.code)
        self.assertNotIn(".VB_VarHelpID", source.code)

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

    def test_excel_guid_userform_export_is_complete(self) -> None:
        text = """VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dialog
   OleObjectBlob   =   \"Dialog.frx\":0000
End
Attribute VB_Name = \"Dialog\"
"""

        self.assertTrue(_is_complete_export("form", text))

    def test_duplicate_component_names_are_rejected(self) -> None:
        with self.assertRaisesRegex(VbaImportError, "Duplicate VBA component"):
            discover_sources(FIXTURES / "duplicates")

    def test_document_module_is_updated_as_code(self) -> None:
        source = parse_source(FIXTURES / "document" / "ThisWorkbook.cls")

        actions = plan_import([source], [ExistingComponent("ThisWorkbook", 100)])

        self.assertEqual(actions[0].operation, "replace-code")

    def test_code_replacement_removes_excel_parentheses_artifact(self) -> None:
        class FakeCodeModule:
            def __init__(self) -> None:
                self.lines = ["old"]

            @property
            def CountOfLines(self) -> int:
                return len(self.lines)

            def DeleteLines(self, start: int, count: int) -> None:
                del self.lines[start - 1 : start - 1 + count]

            def AddFromString(self, code: str) -> None:
                self.lines = code.rstrip().splitlines() + ["", "()"]

            def Lines(self, start: int, count: int) -> str:
                return "\r\n".join(self.lines[start - 1 : start - 1 + count])

        code_module = FakeCodeModule()
        component = SimpleNamespace(Name="InterfaceNTUserName", CodeModule=code_module)

        _replace_component_code(component, "Option Explicit\nSub Run()\nEnd Sub\n")

        self.assertEqual(
            code_module.lines, ["Option Explicit", "Sub Run()", "End Sub", ""]
        )

    def test_code_replacement_accepts_excel_identifier_case_changes(self) -> None:
        class FakeCodeModule:
            def __init__(self) -> None:
                self.lines = ["old"]

            @property
            def CountOfLines(self) -> int:
                return len(self.lines)

            def DeleteLines(self, start: int, count: int) -> None:
                del self.lines[start - 1 : start - 1 + count]

            def AddFromString(self, code: str) -> None:
                self.lines = ["Option Explicit", "sub Run()", "end Sub"]

            def Lines(self, start: int, count: int) -> str:
                return "\r\n".join(self.lines[start - 1 : start - 1 + count])

        component = SimpleNamespace(Name="Example", CodeModule=FakeCodeModule())

        _replace_component_code(component, "Option Explicit\nSub Run()\nEnd Sub\n")

    def test_code_replacement_reports_substantive_excel_changes(self) -> None:
        class FakeCodeModule:
            def __init__(self) -> None:
                self.lines = ["old"]

            @property
            def CountOfLines(self) -> int:
                return len(self.lines)

            def DeleteLines(self, start: int, count: int) -> None:
                del self.lines[start - 1 : start - 1 + count]

            def AddFromString(self, code: str) -> None:
                self.lines = ["Option Explicit", "Sub Other()", "End Function"]

            def Lines(self, start: int, count: int) -> str:
                return "\r\n".join(self.lines[start - 1 : start - 1 + count])

        component = SimpleNamespace(Name="Example", CodeModule=FakeCodeModule())

        with self.assertRaisesRegex(VbaImportError, r"line 2:.*line 3:"):
            _replace_component_code(component, "Option Explicit\nSub Run()\nEnd Sub\n")

    def test_code_replacement_does_not_ignore_string_or_comment_case(self) -> None:
        class FakeCodeModule:
            def __init__(self) -> None:
                self.lines = ["old"]

            @property
            def CountOfLines(self) -> int:
                return len(self.lines)

            def DeleteLines(self, start: int, count: int) -> None:
                del self.lines[start - 1 : start - 1 + count]

            def AddFromString(self, code: str) -> None:
                self.lines = ['MsgBox "HELLO"', "' CHANGED"]

            def Lines(self, start: int, count: int) -> str:
                return "\r\n".join(self.lines[start - 1 : start - 1 + count])

        component = SimpleNamespace(Name="Example", CodeModule=FakeCodeModule())

        with self.assertRaisesRegex(VbaImportError, r"line 1:.*line 2:"):
            _replace_component_code(component, 'MsgBox "hello"\n\' changed\n')

    def test_obsolete_broken_dao_reference_is_removed(self) -> None:
        dao_reference = SimpleNamespace(
            IsBroken=True,
            Guid="{00025E01-0000-0000-C000-000000000046}",
            Name="MISSING: Microsoft DAO 3.6 Object Library",
        )
        excel_reference = SimpleNamespace(
            IsBroken=False,
            Guid="{00020813-0000-0000-C000-000000000046}",
            Name="Excel",
        )

        class FakeReferences:
            def __init__(self) -> None:
                self.items = [excel_reference, dao_reference]

            @property
            def Count(self) -> int:
                return len(self.items)

            def Item(self, index: int) -> object:
                return self.items[index - 1]

            def Remove(self, reference: object) -> None:
                self.items.remove(reference)

        references = FakeReferences()
        _repair_and_validate_references(SimpleNamespace(References=references))

        self.assertEqual(references.items, [excel_reference])

    def test_unknown_broken_reference_stops_import(self) -> None:
        broken_reference = SimpleNamespace(
            IsBroken=True,
            Guid="{11111111-1111-1111-1111-111111111111}",
            Name="MissingLibrary",
        )
        references = SimpleNamespace(
            Count=1,
            Item=lambda index: broken_reference,
            Remove=lambda reference: None,
        )

        with self.assertRaisesRegex(VbaImportError, "MissingLibrary"):
            _repair_and_validate_references(SimpleNamespace(References=references))

    def test_component_type_mismatch_is_rejected(self) -> None:
        source = parse_source(FIXTURES / "module" / "Example.bas")

        with self.assertRaisesRegex(VbaImportError, "type mismatch"):
            plan_import([source], [ExistingComponent("Example", 3)])


if __name__ == "__main__":
    unittest.main()
