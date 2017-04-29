Attribute VB_Name = "Module5"
'matrix algebra
Option Base 1

Public Sub residuals()
' MyArray is a dynamic array of variants.
Dim XArray As Variant, yArray As Variant
Dim zArray As Variant, one As Variant
Dim Fn As Object
Dim a As Variant
Dim y As Variant
Dim x As Variant
Set Fn = Application.WorksheetFunction

'play around with matricess
Dim rng As Range

Set rng = Range("a1").CurrentRegion
ReDim XArray(1 To rng.rows.Count, 1 To rng.Columns.Count)
XArray = rng

Set rng = Range("e1").CurrentRegion
ReDim yArray(1 To rng.rows.Count, 1 To rng.Columns.Count)
yArray = rng

Set rng = Range("a1").CurrentRegion
ReDim XArray(1 To rng.rows.Count, 1 To rng.Columns.Count)
XArray = rng

End Sub
