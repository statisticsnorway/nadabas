import unittest
from pathlib import Path

REPOSITORY_ROOT = Path(__file__).resolve().parents[1]


def procedure(source: str, declaration: str, terminator: str) -> str:
    start = source.index(declaration)
    end = source.index(terminator, start) + len(terminator)
    return source[start:end]


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


class InstallSourceDiscoveryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = (
            REPOSITORY_ROOT / "vba" / "modules" / "InstallMeAsAddin.bas"
        ).read_text(encoding="utf-8-sig")

    def test_installer_prefers_the_canonical_companion_file(self):
        discovery = procedure(
            self.source,
            "Private Function FindSourceNadabasAddIn",
            "End Function",
        )

        self.assertIn("exactPath = installationFolder & NADABAS_ADDIN_NAME", discovery)
        self.assertIn("FindSourceNadabasAddIn = exactPath", discovery)

    def test_installer_searches_versioned_companion_files_to_end_of_loop(self):
        discovery = procedure(
            self.source,
            "Private Function FindSourceNadabasAddIn",
            "End Function",
        )

        self.assertIn('Dir$(installationFolder & "NADABAS*.xlam"', discovery)
        self.assertIn("Do While Len(candidateName) > 0", discovery)
        self.assertIn("candidateName = Dir$()", discovery)
        self.assertIn("More than one NADABAS add-in was found", discovery)

    def test_install_uses_discovered_source_instead_of_a_hard_coded_path(self):
        install = procedure(self.source, "Sub DoInstallAsAddIn", "End Sub")

        self.assertIn("FindSourceNadabasAddIn(ThisWorkbook.path)", install)
        self.assertNotIn('ThisWorkbook.path & "\\NADABAS.xlam"', install)

    def test_managed_loader_is_disabled_after_copy_and_never_deleted(self):
        install = procedure(self.source, "Sub DoInstallAsAddIn", "End Sub")
        managed = procedure(
            install, "If IsManagedNadabasLoader(AI.fullname) Then", "Next AI"
        )
        self.assertNotIn('ElseIf Left$(upperName, 8) = "NADABAS."', install)
        self.assertIn("ElseIf IsNadabasAddInFileName(upperName) Then", install)
        self.assertIn("AI.Installed = False", managed)
        self.assertNotIn("DeleteInstalledNadabasAddIn", managed)
        self.assertLess(install.index("FileCopy sSourceXlam"), install.index(managed))
        self.assertIn("For Each AI In disabledLoaders", install)
        self.assertIn("AI.Installed = True", install)

    def test_managed_loader_detection_uses_the_exact_per_user_paths(self):
        detection = procedure(
            self.source, "Private Function IsManagedNadabasLoader", "End Function"
        )
        self.assertIn('Environ$("LOCALAPPDATA")', detection)
        self.assertIn('managedFolder & "Nadabas.Startup.xll"', detection)
        self.assertIn('managedFolder & "Nadabas.Loader.Test.xll"', detection)
        self.assertEqual(detection.count("StrComp(normalisedCandidate,"), 2)


if __name__ == "__main__":
    unittest.main()
