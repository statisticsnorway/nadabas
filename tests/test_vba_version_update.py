import csv
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

REPOSITORY_ROOT = Path(__file__).resolve().parents[1]


class VbaVersionUpdateTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.update_source = (
            REPOSITORY_ROOT / "vba" / "modules" / "InterfaceVersionUpdate.bas"
        ).read_text(encoding="utf-8-sig")
        cls.startup_source = (
            REPOSITORY_ROOT / "vba" / "modules" / "AutoOpenClose.bas"
        ).read_text(encoding="utf-8-sig")
        cls.open_database_source = (
            REPOSITORY_ROOT / "vba" / "modules" / "cmdOpen_Close_Database.bas"
        ).read_text(encoding="utf-8-sig")
        cls.ribbon_source = (
            REPOSITORY_ROOT / "vba" / "modules" / "RibbonUI.bas"
        ).read_text(encoding="utf-8-sig")
        cls.about_source = (
            REPOSITORY_ROOT / "vba" / "forms" / "dlgAbout.frm"
        ).read_text(encoding="utf-8-sig")
        cls.settings_source = (
            REPOSITORY_ROOT / "vba" / "classes" / "clsUserSettings.cls"
        ).read_text(encoding="utf-8-sig")
        cls.tables_source = (
            REPOSITORY_ROOT / "vba" / "modules" / "DbCreateTables.bas"
        ).read_text(encoding="utf-8-sig")

    def test_release_source_is_version_6_01_003_without_test_override(self):
        self.assertIn(
            'Private Const CURRENT_VERSION As String = "6.01.003"',
            self.update_source,
        )
        self.assertNotIn("TEST_LATEST_VERSION", self.update_source)

    def test_check_is_scheduled_after_startup_preference_is_loaded(self):
        settings_position = self.startup_source.index(
            "InterfaceVersionUpdate.LoadStartupVersionCheckPreference"
        )
        schedule_position = self.startup_source.index(
            "InterfaceVersionUpdate.ScheduleLatestVersionCheck"
        )

        self.assertGreater(schedule_position, settings_position)
        self.assertIn(
            "InterfaceVersionUpdate.CancelScheduledVersionCheck", self.startup_source
        )

    def test_database_setting_is_applied_after_loading(self):
        settings_position = self.open_database_source.index(
            "Usersettings.LoadUserSettings"
        )
        apply_position = self.open_database_source.index(
            "InterfaceVersionUpdate.ApplyDatabaseVersionCheckSetting"
        )

        self.assertGreater(apply_position, settings_position)

    def test_scheduled_check_uses_excel_idle_window_and_weekly_cache(self):
        self.assertIn("Application.OnTime", self.update_source)
        self.assertIn("STARTUP_DELAY_SECONDS As Long = 5", self.update_source)
        self.assertIn("Public Sub RunScheduledVersionCheck()", self.update_source)
        self.assertIn("CHECK_INTERVAL_DAYS As Long = 7", self.update_source)
        self.assertIn("CLng(Date) - lastCheckDay >= CHECK_INTERVAL_DAYS", self.update_source)

    def test_check_uses_latest_published_github_release(self):
        self.assertIn(
            '"https://api.github.com/repos/statisticsnorway/nadabas/releases/latest"',
            self.update_source,
        )
        self.assertIn('request.Open "GET", RELEASE_API_URL, False', self.update_source)
        self.assertIn('ExtractJsonString(responseText, "tag_name")', self.update_source)
        self.assertIn('request.SetRequestHeader "User-Agent"', self.update_source)

    def test_check_is_bounded_and_failure_tolerant(self):
        self.assertIn("request.SetTimeouts 3000, 3000, 5000, 5000", self.update_source)
        self.assertIn('REGISTRY_LAST_CHECK As String = "LastCheckDay"', self.update_source)
        self.assertIn(
            'REGISTRY_LATEST_VERSION As String = "LatestVersion"', self.update_source
        )
        self.assertIn("On Error GoTo CheckFailed", self.update_source)
        self.assertIn("CheckFailed:", self.update_source)

    def test_notification_is_localized_but_never_downloads_or_installs(self):
        self.assertIn("Private Sub ShowUpdateAvailable", self.update_source)
        self.assertIn("MsgBox(messageText, vbYesNo + vbInformation", self.update_source)
        self.assertIn("ThisWorkbook.FollowHyperlink", self.update_source)
        self.assertIn(
            '"https://nadabas.net/nadabas/documents-and-downloads"',
            self.update_source,
        )
        self.assertNotIn("URLDownloadToFile", self.update_source)
        self.assertNotIn("ADODB.Stream", self.update_source)
        self.assertNotIn('request.Open "POST"', self.update_source)

    def test_about_dialog_is_not_used_as_an_update_control(self):
        self.assertIn("Translateform Me", self.about_source)
        self.assertNotIn("InterfaceVersionUpdate", self.about_source)
        self.assertNotIn("WithEvents releaseButton", self.about_source)

    def test_only_administrator_can_change_database_wide_setting(self):
        self.assertIn("Public CheckForUpdates As Boolean", self.settings_source)
        self.assertIn("CheckForUpdates = True", self.settings_source)
        self.assertIn('GetColumn("CheckForUpdates")', self.settings_source)
        self.assertIn('PutColumn "CheckForUpdates"', self.settings_source)
        self.assertIn('AddLongCol "CheckForUpdates"', self.tables_source)
        self.assertIn(
            "If NadabasIsSleeping Or Not isAdministrator Then Exit Sub",
            self.ribbon_source,
        )
        self.assertIn("Usersettings.SaveSettingsToDB", self.ribbon_source)

    def test_both_ribbon_variants_put_update_controls_under_administration(self):
        for name in ("customUI.xml", "customUI14.xml"):
            tree = ET.parse(REPOSITORY_ROOT / "vba" / "customUI" / name)
            parent = {
                child: element for element in tree.iter() for child in element
            }
            controls = {
                element.attrib.get("id"): element
                for element in tree.iter()
                if element.attrib.get("id")
            }

            check_button = controls["btnCheckNadabasVersion"]
            toggle = controls["tglVersionCheck"]
            self.assertEqual(parent[check_button].attrib["id"], "mnuAdministration")
            self.assertEqual(parent[toggle].attrib["id"], "mnuAdministration")
            self.assertEqual(
                check_button.attrib["onAction"],
                "NADABAS.InterfaceVersionUpdate.CheckLatestVersionClick",
            )
            self.assertEqual(
                toggle.attrib["onAction"], "NADABAS.RibbonUI.VersionCheckClick"
            )
            self.assertEqual(
                controls["btnAbout"].attrib["getLabel"],
                "NADABAS.RibbonUI.CommonGetLabel",
            )

    def test_language_resource_rows_match_signed_release(self):
        path = REPOSITORY_ROOT / "vba" / "resources" / "office-ui.csv"
        with path.open(encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.DictReader(handle))

        self.assertEqual(
            {(row["sheet"], row["key"], row["control"]) for row in rows},
            {
                ("Ribbon", "btnCheckNadabasVersion", ""),
                ("Ribbon", "tglVersionCheck", ""),
            },
        )
        for row in rows:
            for language in ("english", "french", "portuguese"):
                self.assertTrue(row[language].strip())


if __name__ == "__main__":
    unittest.main()
