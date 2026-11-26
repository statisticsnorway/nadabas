from pathlib import Path

VBA_MODULES = Path(__file__).resolve().parents[1] / "vba" / "modules"


def test_exchange_uses_its_own_name_in_workbooks_with_both_link_types():
    globals_source = (VBA_MODULES / "Globals.bas").read_text(encoding="ascii")
    lookup = globals_source.split("Public Function GetDescriptionName", 1)[1].split(
        "End Function", 1
    )[0]

    assert (
        lookup.index('awb.Names("ExchDescriptions")')
        < lookup.index("If testDBLinksRange(awb) Then Exit Function")
        < lookup.index('awb.Names("Descriptions")')
    )


def test_validation_checks_and_resizes_the_description_for_the_active_link_type():
    scanner = (VBA_MODULES / "ScanTableDefTest.bas").read_text(encoding="ascii")

    assert "setDescriptionRange(awb, CurrentDB.DbIsExch)" in scanner
    assert "Not DescriptionRange.Worksheet Is DBLinksRange.Worksheet" in scanner
    assert "adjustDescriptionrange awb, DBLinksRange, CurrentDB.DbIsExch" in scanner


def test_inserting_link_rows_and_cleaning_up_preserve_separate_description_names():
    design = (VBA_MODULES / "cmdDesign.bas").read_text(encoding="ascii")

    assert "adjustDescriptionrange awb, GetDBLinksRange(awb), False" in design
    assert "adjustDescriptionrange awb, getExchLinksRange(awb), True" in design
    assert "Set Definition = GetDescriptionName(awb, ForExchange)" in design
    assert 'Areanames.Add "Descriptions"' in design
    assert 'Areanames.Add "ExchDescriptions"' in design
