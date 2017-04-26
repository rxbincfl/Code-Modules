Attribute VB_Name = "Module1"
Sub MyTextToColumns(rng As Variant)
    '
    ' Macro1 Macro
    '
    Dim r As Range
    Sheets("Parameters").Select
    Set r = Range("txt2col")
    r.TextToColumns Destination:=rng, _
    DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, _
    ConsecutiveDelimiter:=True, _
    Tab:=True, _
    Semicolon:=False, _
    Comma:=True, _
    Space:=True, _
    Other:=True, OtherChar:="x", _
    FieldInfo:=Array(Array(1, 9), Array(2, 9), _
    Array(3, 1), Array(4, 9), Array(5, 1)), _
    TrailingMinusNumbers:=True
End Sub
