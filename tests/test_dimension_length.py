from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def procedure(source, declaration, end):
    return source.split(declaration, 1)[1].split(end, 1)[0]


def test_creation_checks_all_lengths_before_schema_mutation():
    source = (ROOT / "vba/modules/DbCreateTables.bas").read_text(encoding="utf-8")
    create = procedure(source, "Public Function CreateNewKeyFam(", "End Function")
    assert create.index("TryDimensionLength(x.Length") < create.index(
        "PrepareTable FName"
    )
    assert create.index("err.Raise") < create.index("PrepareTable FName")
    assert create.index("Next x") < create.index("PrepareTable FName")


def test_form_checks_add_and_finish_before_long_conversion_or_creation():
    source = (ROOT / "vba/forms/frmCreateKeyFamily.frm").read_text(encoding="utf-8")
    add = procedure(source, "Private Sub cmdAdd_Click()", "End Sub")
    finish = procedure(source, "Private Sub cmdFinish_Click()", "End Sub")
    assert add.index("TryDimensionLength") < add.index("ListBox1.AddItem")
    assert finish.index("TryDimensionLength") < finish.index(
        "FN.Length = DimensionLength"
    )
    assert finish.index("TryDimensionLength") < finish.index("If KeyFamilyCreate(")
    create = procedure(source, "Private Function KeyFamilyCreate(", "End Function")
    assert create.index("CreateNewKeyFam(") < create.index("CurrentDB.AddKeyname")
    for body in (add, finish):
        assert 'GetMsg("DimensionLengthInvalid")' in body
