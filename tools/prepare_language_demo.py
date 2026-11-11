"""Create an offline demo from the signed XLAM audit and next-build UI resources."""

import csv
import hashlib
import json
from pathlib import Path
from posixpath import join
from posixpath import normpath
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from test_demo_languages import ADDIN
from test_demo_languages import EXPECTED
from test_demo_languages import ROOT


def prepare():
    if hashlib.sha256(ADDIN.read_bytes()).hexdigest().upper() != EXPECTED:
        raise ValueError("Unexpected demo add-in")
    output = ROOT / "artifacts/demo-language"
    live = json.loads(
        (output / "signed-excel-language-test.json").read_text(encoding="utf-8")
    )
    if live["sha256"] != EXPECTED or not live["originalPreferenceRestored"]:
        raise ValueError("Live evidence for this add-in is required")
    audit = {}
    with ZipFile(ADDIN) as z:
        rel = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        paths = {
            r.get("Id"): normpath(join("xl", r.get("Target"))).lstrip("/") for r in rel
        }
        strings = [
            "".join(e.text or "" for e in si.findall(".//{*}t"))
            for si in ET.fromstring(z.read("xl/sharedStrings.xml"))
        ]
        book = ET.fromstring(z.read("xl/workbook.xml"))
        for sheet in book.findall(".//{*}sheet"):
            name = sheet.get("name")
            if name.lower() not in ("ribbon", "forms", "messages", "errmsg"):
                continue
            path = paths[next(v for k, v in sheet.attrib.items() if k.endswith("}id"))]
            rows = []
            for row in ET.fromstring(z.read(path)).findall(".//{*}row"):
                cells = {}
                for cell in row.findall("{*}c"):
                    value = cell.findtext("{*}v", "")
                    cells[cell.get("r").rstrip("0123456789")] = (
                        strings[int(value)] if cell.get("t") == "s" and value else value
                    )
                if cells.get("A", "").strip() and cells.get("C", "").strip():
                    rows.append(cells)
            audit[name] = {
                "englishTextRows": len(rows),
                "missing": {
                    language: [
                        r.get("A") + "/" + r.get("B", "")
                        for r in rows
                        if not r.get(col, "").strip()
                    ]
                    for language, col in (("fr", "D"), ("pt", "E"), ("id", "F"))
                },
            }
            if name.lower() == "ribbon":
                by_id = {r["A"]: r for r in rows}
                for code, values in live["rows"].items():
                    for key, actual in values.items():
                        if actual != by_id[key]["CDEF"[int(code) - 3]]:
                            raise ValueError(
                                "Resource and Excel runtime result disagree"
                            )
    (output / "signed-resource-coverage.json").write_text(
        json.dumps(audit, indent=2), encoding="utf-8"
    )
    text = json.loads(
        (ROOT / "dotnet/src/Nadabas.ExcelBootstrap/update-text.json").read_text(
            encoding="utf-8"
        )
    )
    columns = ["english", "french", "portuguese", "indonesian"]
    with (ROOT / "vba/resources/office-ui.csv").open(encoding="utf-8-sig") as f:
        exchange = next(
            r for r in csv.DictReader(f) if r["key"] == "tglShowExchangeTools"
        )
    data = {
        c: {"signed": v, "ui": text[c], "exchange": exchange[columns[int(c) - 3]]}
        for c, v in live["rows"].items()
    }
    template = (ROOT / "tools/demo_language_preview.html").read_text(encoding="utf-8")
    (output / "Sprakdemo.html").write_text(
        template.replace(
            "/*DATA*/", json.dumps(data, ensure_ascii=False).replace("<", "\\u003c")
        ),
        encoding="utf-8",
    )
    print(json.dumps(audit))


if __name__ == "__main__":
    prepare()
