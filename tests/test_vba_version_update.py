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
        ).read_text(encoding="ascii")
        cls.startup_source = (
            REPOSITORY_ROOT / "vba" / "modules" / "AutoOpenClose.bas"
        ).read_text(encoding="ascii")
        cls.open_database_source = (
            REPOSITORY_ROOT / "vba" / "modules" / "cmdOpen_Close_Database.bas"
        ).read_text(encoding="cp1252")
        cls.ribbon_source = (
            REPOSITORY_ROOT / "vba" / "modules" / "RibbonUI.bas"
        ).read_text(encoding="cp1252")
        cls.about_source = (
            REPOSITORY_ROOT / "vba" / "forms" / "dlgAbout.frm"
        ).read_text(encoding="cp1252")
        cls.settings_source = (
            REPOSITORY_ROOT / "vba" / "classes" / "clsUserSettings.cls"
        ).read_text(encoding="cp1252")
        cls.tables_source = (
            REPOSITORY_ROOT / "vba" / "modules" / "DbCreateTables.bas"
        ).read_text(encoding="cp1252")

    def test_check_is_scheduled_after_database_settings_are_loaded(self):
        settings_position = self.open_database_source.index(
            "Usersettings.LoadUserSettings"
        )
        schedule_position = self.open_database_source.index(
            "InterfaceVersionUpdate.ScheduleLatestVersionCheck"
        )

        self.assertGreater(schedule_position, settings_position)
        self.assertNotIn(
            "InterfaceVersionUpdate.CheckForLatestVersion", self.startup_source
        )
        self.assertIn(
            "InterfaceVersionUpdate.CancelScheduledVersionCheck", self.startup_source
        )

    def test_scheduled_check_uses_excel_idle_window(self):
        self.assertIn("Application.OnTime", self.update_source)
        self.assertIn("SCHEDULE_DELAY_SECONDS As Long = 5", self.update_source)
        self.assertIn("SCHEDULE_WINDOW_SECONDS As Long = 30", self.update_source)
        self.assertIn("Public Sub RunScheduledVersionCheck()", self.update_source)

    def test_check_uses_the_latest_published_github_release(self):
        self.assertIn(
            '"https://api.github.com/repos/statisticsnorway/nadabas/releases/latest"',
            self.update_source,
        )
        self.assertIn(
            'request.Open "GET", LATEST_RELEASE_API, False', self.update_source
        )
        self.assertIn('"tag_name"', self.update_source)
        self.assertIn(
            'request.SetRequestHeader "X-GitHub-Api-Version"', self.update_source
        )
        self.assertIn('request.SetRequestHeader "User-Agent"', self.update_source)

    def test_check_is_bounded_cached_and_failure_tolerant(self):
        self.assertIn("request.SetTimeouts 1000, 1500, 1500, 2500", self.update_source)
        self.assertIn(
            'REGISTRY_LAST_CHECK As String = "LastCheckDay"', self.update_source
        )
        self.assertIn(
            'REGISTRY_LATEST_VERSION As String = "LatestVersion"', self.update_source
        )
        self.assertIn("If Not UpdateCheckIsDue Then Exit Sub", self.update_source)
        self.assertIn("CHECK_INTERVAL_DAYS As Long = 7", self.update_source)
        self.assertIn(
            "currentDay - lastCheckDay >= CHECK_INTERVAL_DAYS",
            self.update_source,
        )
        self.assertIn("On Error GoTo CheckFailed", self.update_source)
        self.assertIn("CheckFailed:", self.update_source)

    def test_notification_is_non_modal_and_does_not_install(self):
        self.assertNotIn("MsgBox", self.update_source)
        self.assertNotIn("URLDownloadToFile", self.update_source)
        self.assertNotIn("ADODB.Stream", self.update_source)
        self.assertNotIn('request.Open "POST"', self.update_source)
        self.assertIn("ThisWorkbook.FollowHyperlink", self.update_source)

    def test_newer_equal_older_and_invalid_cases_have_vba_self_tests(self):
        self.assertIn(
            'Debug.Assert CompareVersions("6.01.003", "6.01.002") = 1',
            self.update_source,
        )
        self.assertIn(
            'Debug.Assert CompareVersions("6.01.002", "6.01.002") = 0',
            self.update_source,
        )
        self.assertIn(
            'Debug.Assert CompareVersions("6.01.001", "6.01.002") = -1',
            self.update_source,
        )
        self.assertIn(
            'Debug.Assert NormalizeVersion("not-a-version") = ""',
            self.update_source,
        )

    def test_about_dialog_owns_the_release_details(self):
        self.assertIn("Public Sub AboutGetLabel", self.ribbon_source)
        self.assertIn('GetRibbon("btnAboutUpdate")', self.ribbon_source)
        self.assertIn("Private WithEvents releaseButton", self.about_source)
        self.assertIn('"lblVersionUpdateAvailable"', self.about_source)
        self.assertIn('"lblLatestRelease"', self.about_source)
        self.assertIn('"cmdVersionRelease"', self.about_source)
        self.assertIn("InterfaceVersionUpdate.OpenDownloadPage", self.about_source)
        self.assertIn(
            '"https://nadabas.net/nadabas/documents-and-downloads"',
            self.update_source,
        )

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

    def test_both_ribbon_variants_use_dynamic_about_and_admin_toggle(self):
        for name in ("customUI.xml", "customUI14.xml"):
            path = REPOSITORY_ROOT / "vba" / "customUI" / name
            tree = ET.parse(path)
            controls = {
                element.attrib.get("id"): element.attrib
                for element in tree.iter()
                if element.attrib.get("id")
            }
            self.assertEqual(
                controls["btnAbout"]["getLabel"],
                "NADABAS.RibbonUI.AboutGetLabel",
            )
            self.assertEqual(
                controls["tglVersionCheck"]["onAction"],
                "NADABAS.RibbonUI.VersionCheckClick",
            )

    def test_language_resource_rows_cover_ribbon_and_about_dialog(self):
        path = REPOSITORY_ROOT / "vba" / "resources" / "office-ui.csv"
        with path.open(encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.DictReader(handle))
        identities = {(row["sheet"], row["key"], row["control"]) for row in rows}

        self.assertIn(("Ribbon", "btnAboutUpdate", ""), identities)
        self.assertIn(("Ribbon", "tglVersionCheck", ""), identities)
        self.assertIn(("Forms", "dlgAbout", "lblVersionUpdateAvailable"), identities)
        self.assertIn(("Forms", "dlgAbout", "lblLatestRelease"), identities)
        self.assertIn(("Forms", "dlgAbout", "cmdVersionRelease"), identities)


if __name__ == "__main__":
    unittest.main()
