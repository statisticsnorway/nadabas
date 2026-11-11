import unittest

from tools.prepare_upgrade_test import stamp_build_date


class BuildDateTests(unittest.TestCase):
    def test_stamp_is_fixed_and_accepts_vba_identifier_case(self):
        code, form = stamp_build_date(
            'Private Const CURRENT_VERSION As String = "6.02.008"',
            "Translateform Me\nMe.VersionNumber.Caption = InterfaceVersionUpdate.installedVersion",
            "2026-11-03",
        )
        self.assertIn('RELEASE_DATE_ISO As String = "2026-11-03"', code)
        self.assertNotIn("Format$(Date", code)
        self.assertIn("Me.Label2.Caption = InterfaceVersionUpdate.AboutBuildDate", form)
        self.assertLess(form.index("Translateform"), form.index("Me.Label2"))

    def test_invalid_date_or_unknown_form_is_rejected(self):
        with self.assertRaises(ValueError):
            stamp_build_date("", "", "2026-02-30")
        with self.assertRaises(RuntimeError):
            stamp_build_date(
                'Private Const CURRENT_VERSION As String = "1"', "unknown", "2026-11-03"
            )

    def test_already_stamped_source_is_rejected(self):
        with self.assertRaises(RuntimeError):
            stamp_build_date(
                'Private Const RELEASE_DATE_ISO As String = "2026-11-03"',
                "",
                "2026-11-04",
            )
