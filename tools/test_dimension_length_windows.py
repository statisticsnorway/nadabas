"""Run the production VBA length parser in an owned, unsaved Excel workbook.

Requires Excel and existing trusted VBA-project access. Does not change that
setting or open a database. Only the scratch workbook and its Excel instance
are closed. Run explicitly; this is not part of unattended pytest.
"""

import json
from pathlib import Path


def main():
    import win32com.client

    source = (
        Path(__file__).resolve().parents[1] / "vba/modules/DbCreateTables.bas"
    ).read_text(encoding="utf-8")
    declaration = "Public Function TryDimensionLength("
    parser = (
        declaration
        + source.split(declaration, 1)[1].split("End Function", 1)[0]
        + "End Function"
    )
    excel = win32com.client.DispatchEx("Excel.Application")
    book = None
    try:
        excel.EnableEvents = False
        book = excel.Workbooks.Add()
        module = book.VBProject.VBComponents.Add(1)
        module.Name = "DimensionLengthProbe"
        module.CodeModule.AddFromString("Option Explicit\n" + parser + """
Public Function Probe(ByVal InputValue As Variant) As Long
Dim Length As Long
If TryDimensionLength(InputValue, Length) Then Probe = Length
End Function
""")
        cases = [
            ("1", 1),
            ("255", 255),
            (" 20 ", 20),
            ("256", 0),
            ("0", 0),
            ("-1", 0),
            ("", 0),
            ("abc", 0),
            ("99999999999999999999999", 0),
            ("1E9999", 0),
            (1.5, 0),
            (254.5, 0),
            (255.1, 0),
            (True, 0),
        ]
        for value, expected in cases:
            actual = excel.Run(f"'{book.Name}'!DimensionLengthProbe.Probe", value)
            assert actual == expected, (value, expected, actual)
        print(json.dumps({"status": "passed", "vbaParserCases": len(cases)}))
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        excel.Quit()


if __name__ == "__main__":
    main()
