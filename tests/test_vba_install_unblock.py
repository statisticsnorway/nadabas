import csv
import unittest
from pathlib import Path


REPOSITORY_ROOT = Path(__file__).resolve().parents[1]


class VbaInstallUnblockTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.installer_source = (
            REPOSITORY_ROOT / "vba" / "modules" / "InstallMeAsAddin.bas"
        ).read_text(encoding="ascii")
        resource_path = (
            REPOSITORY_ROOT / "vba" / "resources" / "office-ui.csv"
        )
        with resource_path.open(encoding="utf-8-sig", newline="") as handle:
            cls.resource_rows = list(csv.DictReader(handle))

    def test_file_not_found_has_localized_unblock_guidance(self):
        self.assertIn("installErrorNumber = err.Number", self.installer_source)
        self.assertIn("If installErrorNumber = 53 Then", self.installer_source)
        for message_id in ("M215A", "M215B", "M215C"):
            self.assertIn(f'"{message_id}"', self.installer_source)
        self.assertIn("Private Function InstallMessage", self.installer_source)

    def test_other_install_errors_remain_localized_and_include_details(self):
        self.assertIn(
            '"M216", installErrorDescription', self.installer_source
        )

    def test_all_installer_messages_have_three_languages(self):
        rows = {
            row["key"]: row
            for row in self.resource_rows
            if row["sheet"] == "Messages"
        }
        self.assertEqual(set(rows), {"M215A", "M215B", "M215C", "M216"})
        for row in rows.values():
            for language in ("english", "french", "portuguese"):
                self.assertTrue(row[language].strip())

    def test_guidance_targets_the_extracted_file_not_the_zip(self):
        rows = {
            row["key"]: row
            for row in self.resource_rows
            if row["sheet"] == "Messages"
        }
        for language in ("english", "french", "portuguese"):
            combined = " ".join(rows[key][language] for key in ("M215A", "M215B", "M215C"))
            self.assertIn("nadabas.xlam", combined.casefold())
            self.assertNotIn("zip", combined.casefold())


if __name__ == "__main__":
    unittest.main()
