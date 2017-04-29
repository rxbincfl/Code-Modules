Attribute VB_Name = "UpdatePQ"
Option Explicit

Const StockFileName = "Stock Price File"
Dim fn As WorksheetFunction
Dim myTable As ListObject
Public myArray As Variant
Public rng As Variant
Public wb As Workbook

Public Sub UpdatePowerQueries()
' Macro to update my Power Query script(s)

Dim lTest As Long, cn As WorkbookConnection

On Error Resume Next
For Each cn In ThisWorkbook.Connections
    lTest = InStr(1, cn.OLEDBConnection.Connection, "Provider=Microsoft.Mashup.OleDb.1", vbTextCompare)
        If Err.Number <> 0 Then
            Err.Clear
            'Exit For
        End If
'==============================================================='
    'do we need error code here if cn does not exist?
    If lTest > 0 Then cn.Refresh
Next cn

End Sub


Function Age(DoB As Date)
    If DoB = 0 Then
        Age = "No Birthdate"
    Else
        Select Case Month(Date)
            Case Is < Month(DoB)
                Age = Year(Date) - Year(DoB) - 1
            Case Is = Month(DoB)
                If Day(Date) >= Day(DoB) Then
                    Age = Year(Date) - Year(DoB)
                Else
                    Age = Year(Date) - Year(DoB) - 1
                End If
            Case Is > Month(DoB)
                Age = Year(Date) - Year(DoB)
        End Select
    End If
End Function

