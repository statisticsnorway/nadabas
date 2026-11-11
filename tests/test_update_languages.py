import json
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]


class UpdateLanguagesTests(unittest.TestCase):
    def test_complete_catalog_and_ready_version_in_each_language(self):
        tables = json.loads(
            (ROOT / "dotnet/src/Nadabas.ExcelBootstrap/update-text.json").read_text(
                encoding="utf-8"
            )
        )
        self.assertEqual(set(tables), {"3", "4", "5", "6"})
        for code, table in tables.items():
            self.assertEqual(set(table), set(tables["3"]))
            self.assertTrue(all(text.strip() for text in table.values()))
            self.assertIn("6.02.008", table["ready"].format("6.02.008"))
            if code != "3":
                self.assertNotEqual(table["restart"], tables["3"]["restart"])

    def test_exchange_toggle_uses_translatable_callbacks_in_both_schemas(self):
        for filename in ("customUI.xml", "customUI14.xml"):
            root = ET.parse(ROOT / "vba/customUI" / filename)
            toggle = next(
                x for x in root.iter() if x.get("id") == "tglShowExchangeTools"
            )
            self.assertEqual(toggle.get("getLabel"), "NADABAS.RibbonUI.CommonGetLabel")
            self.assertEqual(
                toggle.get("getScreentip"), "NADABAS.RibbonUI.ExchangeToolsTip"
            )
            self.assertEqual(
                toggle.get("getSupertip"), "NADABAS.RibbonUI.ExchangeToolsHelp"
            )
            self.assertNotIn("label", toggle.attrib)
