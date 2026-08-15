import unittest
from pathlib import Path

REPOSITORY_ROOT = Path(__file__).resolve().parents[1]


class InstallUnblockDocumentationTests(unittest.TestCase):
    def test_operations_guide_explains_manual_unblock_recovery(self):
        documentation = (
            REPOSITORY_ROOT / "docs" / "operations" / "index.qmd"
        ).read_text(encoding="utf-8")
        documentation = " ".join(documentation.split())

        for expected in (
            "run-time error 53",
            "NADABAS.xlam",
            "Properties",
            "Unblock",
            "official NADABAS download channel",
        ):
            self.assertIn(expected, documentation)


if __name__ == "__main__":
    unittest.main()
