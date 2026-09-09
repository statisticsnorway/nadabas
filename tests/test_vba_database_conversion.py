from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CONVERSION = (ROOT / "vba" / "modules" / "cmdConvertDB.bas").read_text(
    encoding="utf-8-sig"
)
DB_INTERFACE = (ROOT / "vba" / "modules" / "DBInterface.bas").read_text(
    encoding="utf-8-sig"
)


def test_key_family_dimensions_are_selected_by_position_before_value():
    dimension_loop = CONVERSION[
        CONVERSION.index("' Build list of dimension fields.") : CONVERSION.index(
            "' The Value field is required"
        )
    ]

    assert "If ValueFN Is Nothing Then" in dimension_loop
    assert 'UCase$(Trim$(FN.name)) = "VALUE"' in dimension_loop
    assert "xclsFieldNames.Add FN" in dimension_loop
    assert "IsStandardDataColumn" not in dimension_loop


def test_status_is_not_rejected_as_a_dimension_name():
    assert "Private Function IsStandardDataColumn" not in CONVERSION
    assert '"STATUS", _' not in CONVERSION


def test_copy_preflights_target_columns_before_copying_rows():
    schema_check = CONVERSION.index("If Not CopyCursorSchemasAreCompatible(sTable)")
    row_loop = CONVERSION.index("Do While CursorGetEoF = False", schema_check)

    assert schema_check < row_loop
    assert "Missing SQL column(s): " in CONVERSION
    assert "No rows were copied for this table." in CONVERSION


def test_conversion_uses_column_write_that_propagates_errors():
    assert "PutColumnStrict f.name, f.value" in CONVERSION
    assert "Public Sub PutColumnStrict" in DB_INTERFACE

    strict_method = DB_INTERFACE[
        DB_INTERFACE.index("Public Sub PutColumnStrict") : DB_INTERFACE.index(
            "Public Sub PutColumnNull"
        )
    ]
    assert "On Error" not in strict_method
