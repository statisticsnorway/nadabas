import unittest
import zipfile
from pathlib import Path

from tools.apply_office_ui import OfficeUiError
from tools.apply_office_ui import read_resource_rows
from tools.apply_office_ui import replace_custom_ui

REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
TEST_TEMP_ROOT = REPOSITORY_ROOT / "build" / "test-fixtures"


class ApplyOfficeUiTests(unittest.TestCase):
    def test_repository_resource_rows_are_valid(self):
        rows = read_resource_rows(
            REPOSITORY_ROOT / "vba" / "resources" / "office-ui.csv"
        )
        self.assertEqual(len(rows), 8)
        self.assertEqual({row.sheet for row in rows}, {"Messages", "Ribbon"})

    def test_custom_ui_replacement_preserves_unrelated_package_parts(self):
        TEST_TEMP_ROOT.mkdir(parents=True, exist_ok=True)
        package = TEST_TEMP_ROOT / "test-package.xlam"
        custom_ui = TEST_TEMP_ROOT / "source" / "customUI"
        custom_ui.mkdir(parents=True, exist_ok=True)
        (custom_ui / "customUI.xml").write_text("new-2006", encoding="utf-8")
        (custom_ui / "customUI14.xml").write_text("new-2010", encoding="utf-8")

        with zipfile.ZipFile(package, "w") as archive:
            archive.writestr("customUI/customUI.xml", "old-2006")
            archive.writestr("customUI/customUI14.xml", "old-2010")
            archive.writestr("xl/workbook.xml", "unchanged")

        replace_custom_ui(package, custom_ui)

        with zipfile.ZipFile(package) as archive:
            self.assertEqual(archive.read("customUI/customUI.xml"), b"new-2006")
            self.assertEqual(archive.read("customUI/customUI14.xml"), b"new-2010")
            self.assertEqual(archive.read("xl/workbook.xml"), b"unchanged")

    def test_resource_file_rejects_unsupported_sheet(self):
        TEST_TEMP_ROOT.mkdir(parents=True, exist_ok=True)
        path = TEST_TEMP_ROOT / "invalid-resources.csv"
        path.write_text(
            "sheet,key,control,english,french,portuguese\n"
            "Other,key,,English,French,Portuguese\n",
            encoding="utf-8",
        )
        with self.assertRaises(OfficeUiError):
            read_resource_rows(path)


if __name__ == "__main__":
    unittest.main()
