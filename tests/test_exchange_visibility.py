import re
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

from tools.exchange_ui_patch import BRIDGE
from tools.exchange_ui_patch import GUARDS
from tools.exchange_ui_patch import patch_code
from tools.exchange_ui_patch import patch_xml

ROOT = Path(__file__).resolve().parents[1]


class ExchangeVisibilityTests(unittest.TestCase):
    def test_both_ribbons_keep_toggle_accessible_and_gate_all_exchange_groups(self):
        source = (ROOT / "vba/modules/RibbonUI.bas").read_text(encoding="utf-8-sig")
        for name in ("customUI.xml", "customUI14.xml"):
            tree = ET.parse(ROOT / "vba/customUI" / name)
            parents = {child: parent for parent in tree.iter() for child in parent}
            controls = {e.attrib.get("id"): e for e in tree.iter()}
            toggle = controls["tglShowExchangeTools"]
            self.assertEqual(parents[toggle].attrib["id"], "mnuAdministration")
            self.assertIs(parents[toggle], parents[controls["btnCreateExchLink"]])
            siblings = list(parents[toggle])
            self.assertEqual(
                siblings.index(toggle),
                siblings.index(controls["btnCreateExchLink"]) + 1,
            )
            self.assertEqual(
                toggle.attrib["getVisible"], "NADABAS.RibbonUI.ExchangeToolsVisible"
            )
            self.assertEqual(
                toggle.attrib["getPressed"], "NADABAS.RibbonUI.ExchangeToolsPressed"
            )
            for group in (
                "grpExpLoadSave",
                "grpExchDatabase",
                "grpExchClassifications",
                "grpExchDesign",
                "grpExchAdministration",
            ):
                callback = controls[group].attrib["getVisible"].split(".")[-1]
                self.assertIn(callback, GUARDS)
                procedure = re.search(
                    r"Sub " + callback + r"\(.*?End Sub", source, re.S
                )[0]
                self.assertLess(
                    procedure.index("If Not ExchangeToolsShown"),
                    procedure.index("CurrentWBData"),
                )

    def test_preference_defaults_off_and_changes_only_display_setting(self):
        self.assertIn('"ShowExchangeTools", "0"', BRIDGE)
        self.assertIn("DoInvalidateIf", BRIDGE)
        for forbidden in (
            "OpenExchDatabase",
            "CloseExchDatabase",
            "Delete",
            "CurrentDB",
        ):
            self.assertNotIn(forbidden, BRIDGE)

    def test_toggle_visibility_uses_actual_links_independently_of_saved_preference(
        self,
    ):
        callback = re.search(r"Sub ExchangeToolsVisible\(.*?End Sub", BRIDGE, re.S)[0]
        self.assertIn("If ActiveWorkbook Is Nothing Then Exit Sub", callback)
        self.assertIn("TestExchLinksRange(ActiveWorkbook)", callback)
        self.assertNotIn("ExchangeToolsShown", callback)
        self.assertNotIn("CurrentWBData", callback)
        self.assertIn("On Error GoTo Hidden", callback)

    def test_patch_rejects_duplicates_and_unexpected_input(self):
        with self.assertRaises(ValueError):
            patch_code("already ExchangeToolsShown")
        with self.assertRaises(ValueError):
            patch_code("unknown")
        with self.assertRaises(ValueError):
            patch_xml('<group id="other"/>')
        xml = patch_xml(
            '<menu id="mnuAdministration"><button id="btnCreateExchLink"/></menu>'
        )
        self.assertEqual(xml.count('id="tglShowExchangeTools"'), 1)
        with self.assertRaises(ValueError):
            patch_xml(xml)
