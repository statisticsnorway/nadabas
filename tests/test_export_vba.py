import unittest
from pathlib import Path

from tools.export_vba import VBEXT_CT_CLASS_MODULE
from tools.export_vba import VBEXT_CT_DOCUMENT
from tools.export_vba import VBEXT_CT_MS_FORM
from tools.export_vba import VBEXT_CT_STANDARD_MODULE
from tools.export_vba import VbaExportError
from tools.export_vba import _normalise_text_export
from tools.export_vba import component_relative_path


class ExportVbaTests(unittest.TestCase):
    def test_component_layout_matches_repository(self) -> None:
        self.assertEqual(
            component_relative_path("Module1", VBEXT_CT_STANDARD_MODULE),
            Path("modules/Module1.bas"),
        )
        self.assertEqual(
            component_relative_path("Class1", VBEXT_CT_CLASS_MODULE),
            Path("classes/Class1.cls"),
        )
        self.assertEqual(
            component_relative_path("ThisWorkbook", VBEXT_CT_DOCUMENT),
            Path("classes/ThisWorkbook.cls"),
        )
        self.assertEqual(
            component_relative_path("Dialog1", VBEXT_CT_MS_FORM),
            Path("forms/Dialog1.frm"),
        )

    def test_unsafe_component_name_is_rejected(self) -> None:
        with self.assertRaisesRegex(VbaExportError, "Unsafe VBA component name"):
            component_relative_path("../Module1", VBEXT_CT_STANDARD_MODULE)

    def test_unsupported_component_type_is_rejected(self) -> None:
        with self.assertRaisesRegex(VbaExportError, "Unsupported VBA component type"):
            component_relative_path("Module1", 999)

    def test_text_export_is_normalised_for_git(self) -> None:
        path = Path(__file__).with_name(".export-vba-normalise.tmp")
        self.addCleanup(path.unlink, missing_ok=True)
        try:
            path.write_bytes(
                b'Attribute VB_Name = "Module1"  \r\n\r\nSub Run()\t\r\nEnd Sub\r\n \t\r\n'
            )

            _normalise_text_export(path)

            self.assertEqual(
                path.read_bytes(),
                b'Attribute VB_Name = "Module1"\r\n\r\nSub Run()\r\nEnd Sub\r\n',
            )
        finally:
            path.unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
