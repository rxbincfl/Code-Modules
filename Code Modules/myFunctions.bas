Attribute VB_Name = "myFunctions"
Option Explicit

Dim rng As Variant

Public Function fnStockPrice(name As String) As Double
Dim symb As Double
Dim myTable As ListObject
Dim arr As Variant

    On Error Resume Next
    Application.Volatile
    
    'Set path for Table variable
    With Workbooks("stockdata.xlsm").Worksheets("Parameters")
        Set myTable = .ListObjects("closing_prices")
        'debug
        Set arr = myTable
    End With
        
    'Create Array List from Table
    Set rng = myTable.ListColumns(1).DataBodyRange
    
    fnStockPrice = FindXY2(name)
    
End Function

Function FindXY2(symb As String)
Dim j As Long
Dim n As Long
Dim dTime As Double
    
    On Error GoTo Finish
    
    With Application.WorksheetFunction
        FindXY2 = .Index(rng.Offset(0, 1), .Match(symb, rng, 0))
    End With
    Exit Function
Finish:
    If (FindXY2 = 0) Then FindXY2 = 999999
End Function

