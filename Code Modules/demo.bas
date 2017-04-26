Attribute VB_Name = "demo"
Sub demo()
Dim Excel_workbook As Excel.Workbook
'   Set Excel_workbook = Workbooks.Open("age.xlam")
   ' some code goes here
   ' at the end write the below statement
'   Set Excel_workbook = Nothing
   
Dim a, m(), b, c, e As Variant 'default is variant
Set d = New MatrixMath
Dim MyArr() As Variant

a = [{3,5,1; 9,1,4}] '2 rows, 3 cols
b = [{1,1}] '1 row, 2 cols


With WorksheetFunction
    m = .MMult((.MMult(.Transpose(b), b)), a)
'    redim preserve m as Double
End With

Set d = New MatrixMath
ReDim b(0, 0)
b = d.SetArrayA(tmparray, 1)
'm = d.SetArrayA(m, 1)

End Sub


