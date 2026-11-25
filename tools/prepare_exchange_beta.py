"""Prepare a separate beta with Exchange UI control and loader integration."""

import argparse
import hashlib
import json
import re
from datetime import date
from pathlib import Path
from zipfile import ZipFile

from apply_office_ui import _resource_row_number
from apply_office_ui import read_resource_rows
from exchange_ui_patch import patch_code
from exchange_ui_patch import patch_xml
from prepare_upgrade_test import managed_bridge
from prepare_upgrade_test import stamp_build_date

BETA_HASH = "749DDB3F496C6796747A8EDA15BF55C7B57F053A1E50E822029F2AF845824604"


def main():
    import win32com.client

    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("target", type=Path)
    parser.add_argument("--build-date", default=date.today().isoformat())
    args = parser.parse_args()
    source, target = args.source.resolve(), args.target.resolve()
    if hashlib.sha256(source.read_bytes()).hexdigest().upper() != BETA_HASH:
        raise RuntimeError("Not the reviewed original beta")
    if target.exists():
        raise RuntimeError("Output must be new")
    date.fromisoformat(args.build_date)
    target.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(source) as original, ZipFile(target, "x") as output:
        for item in original.infolist():
            data = original.read(item.filename)
            if item.filename in ("customUI/customUI.xml", "customUI/customUI14.xml"):
                data = patch_xml(data.decode("utf-8-sig")).encode("utf-8")
            output.writestr(item, data)
    excel = win32com.client.DispatchEx("Excel.Application")
    book = None
    try:
        excel.AutomationSecurity = 3
        excel.EnableEvents = False
        book = excel.Workbooks.Open(str(target), UpdateLinks=0, ReadOnly=False)
        if book.VBProject.Protection != 0:
            raise RuntimeError("Locked project; no protection changed")
        ribbon = book.VBProject.VBComponents.Item("RibbonUI").CodeModule
        code = patch_code(ribbon.Lines(1, ribbon.CountOfLines))
        ribbon.DeleteLines(1, ribbon.CountOfLines)
        ribbon.AddFromString(code)
        module = book.VBProject.VBComponents.Item("InterfaceVersionUpdate").CodeModule
        about = book.VBProject.VBComponents.Item("dlgAbout").CodeModule
        code, form = stamp_build_date(
            managed_bridge(module.Lines(1, module.CountOfLines)),
            about.Lines(1, about.CountOfLines),
            args.build_date,
        )
        module.DeleteLines(1, module.CountOfLines)
        module.AddFromString(code)
        about.DeleteLines(1, about.CountOfLines)
        about.AddFromString(form)
        resources = Path(__file__).resolve().parents[1] / "vba/resources/office-ui.csv"
        keys = {
            "tglShowExchangeTools",
            "tipExchangeTools",
            "helpExchangeTools",
            "ExchangePreferenceError",
            "ManagedUpdateUnavailable",
        }
        for row in read_resource_rows(resources):
            if row.key in keys:
                sheet = book.Worksheets(row.sheet)
                number = _resource_row_number(sheet, row)
                for column, value in enumerate(
                    (
                        row.key,
                        row.control,
                        row.english,
                        row.french,
                        row.portuguese,
                        row.indonesian,
                    ),
                    1,
                ):
                    sheet.Cells(number, column).Value2 = value
        dialog = book.VBProject.VBComponents.Item("dlgSelectLanguage").CodeModule
        code = dialog.Lines(1, dialog.CountOfLines)
        reviewed = (resources.parents[1] / "forms/dlgSelectLanguage.frm").read_text(
            encoding="utf-8-sig"
        )
        pattern = r"Private Sub UserForm_Initialize\(\).*?End Sub"
        initializer = re.search(pattern, reviewed, re.S | re.I)
        if initializer is None or len(re.findall(pattern, code, re.S | re.I)) != 1:
            raise RuntimeError("Unexpected language selection form")
        dialog.DeleteLines(1, dialog.CountOfLines)
        dialog.AddFromString(
            re.sub(pattern, lambda _: initializer[0], code, flags=re.S | re.I)
        )
        book.Save()
    finally:
        if book is not None:
            book.Close(False)
        excel.Quit()
    if hashlib.sha256(source.read_bytes()).hexdigest().upper() != BETA_HASH:
        raise RuntimeError("Original beta changed")
    target.with_suffix(".json").write_text(
        json.dumps(
            {
                "testOnly": True,
                "publishedRelease": False,
                "sourceSha256": BETA_HASH,
                "releaseVersion": "6.02.006",
                "buildDate": args.build_date,
                "changes": [
                    "Exchange visibility preference, default off",
                    "Managed update bridge",
                    "Fixed About build date",
                ],
                "signing": "must-sign-and-verify-before-packaging",
            },
            indent=2,
        )
        + "\n"
    )
    print(target)


if __name__ == "__main__":
    main()
