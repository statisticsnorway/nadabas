"""Narrow, reviewable Exchange ribbon visibility patch; no database changes."""

import re

GUARDS = ("VisibleExchange", "VisibleExchDesign", "grpExchAdministrationVisible")
BRIDGE = """
' BEGIN EXCHANGE VISIBILITY PREFERENCE
Private Function ExchangeToolsShown() As Boolean
    On Error GoTo Hidden
    ExchangeToolsShown = (GetSetting("NADABAS", "Ribbon", "ShowExchangeTools", "0") = "1")
    Exit Function
Hidden:
    ExchangeToolsShown = False
End Function

Public Sub ExchangeToolsVisible(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo Hidden
    returnedVal = False
    If ActiveWorkbook Is Nothing Then Exit Sub
    returnedVal = TestExchLinksRange(ActiveWorkbook)
    Exit Sub
Hidden:
    returnedVal = False
End Sub

Public Sub ExchangeToolsTip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = MultiLang.GetRibbon("tipExchangeTools")
End Sub

Public Sub ExchangeToolsHelp(control As IRibbonControl, ByRef returnedVal)
    returnedVal = MultiLang.GetRibbon("helpExchangeTools")
End Sub

Public Sub ExchangeToolsPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = ExchangeToolsShown
End Sub

Public Sub ExchangeToolsClick(control As IRibbonControl, pressed As Boolean)
    On Error GoTo SettingFailed
    SaveSetting "NADABAS", "Ribbon", "ShowExchangeTools", IIf(pressed, "1", "0")
    DoInvalidateIf
    Exit Sub
SettingFailed:
    MsgBox MultiLang.GetMsg("ExchangePreferenceError"), vbExclamation, "NADABAS"
    DoInvalidateIf
End Sub
' END EXCHANGE VISIBILITY PREFERENCE
"""

TOGGLE = """
              <toggleButton id="tglShowExchangeTools"
                      getLabel="NADABAS.RibbonUI.CommonGetLabel"
                      getScreentip="NADABAS.RibbonUI.ExchangeToolsTip"
                      getSupertip="NADABAS.RibbonUI.ExchangeToolsHelp"
                      getVisible="NADABAS.RibbonUI.ExchangeToolsVisible"
                      getPressed="NADABAS.RibbonUI.ExchangeToolsPressed"
                      onAction="NADABAS.RibbonUI.ExchangeToolsClick" />
"""


def patch_code(source):
    if "ExchangeToolsShown" in source:
        raise ValueError("Exchange visibility patch already present")
    for name in GUARDS:
        pattern = rf"((?:Public )?Sub {name}\([^\n]*\n.*?returnedVal = False[^\n]*\n)"
        source, count = re.subn(
            pattern,
            lambda m: m[1] + "    If Not ExchangeToolsShown Then Exit Sub\n",
            source,
            flags=re.DOTALL | re.IGNORECASE,
        )
        if count != 1:
            raise ValueError("Unexpected visibility callback: " + name)
    return source.rstrip() + "\n" + BRIDGE


def patch_xml(source):
    if "tglShowExchangeTools" in source:
        raise ValueError("Exchange toggle already present")
    pattern = r'(<group\s+id="grpAbout"[^>]*>.*?)(</group>)'
    source, count = re.subn(
        pattern, lambda m: m[1] + TOGGLE + m[2], source, flags=re.DOTALL
    )
    if count != 1:
        raise ValueError("Expected one always-visible About group")
    return source
