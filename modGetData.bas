Attribute VB_Name = "modGetData"
Public Sub GetData()

Dim QuerySheet As Worksheet
Dim DataSheet As Worksheet
Dim qurl As String
Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim xObject As Object


    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
'    Application.Calculation = xlCalculationManual
    
    'sheet code
    Set DataSheet = Sheet1
    
'    Range("t_code").Offset(0, 2).CurrentRegion.ClearContents
    'make sure col B is empty
    ' clear all previous connections
    For Each xObject In ActiveWorkbook.Connections
        xObject.Delete
    Next
    
    'get the names for downloading
    i = Range("t_code").Row
    qurl = "http://download.finance.yahoo.com/d/quotes.csv?s="
    For Each Row In Range("t_code").Rows
        qurl = qurl + "+" + Row
    Next
    
    qurl = qurl + "&f=" + Range("C2")
    Range("c1") = qurl
    
QueryQuote:
    With ActiveSheet.QueryTables.Add(Connection:="URL;" & qurl, _
        Destination:=DataSheet.Range("t_code").Offset(0, 1))
        .BackgroundQuery = True
        .TablesOnlyFromHTML = False
        .Refresh BackgroundQuery:=False
        .SaveData = True
    End With
    

'    Range("t_code").Offset(0, 1).CurrentRegion.TextToColumns Destination:=Range("t_code").Offset(0, 2), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=True, Space:=False, Other:=False
        

    'turn calculation back on
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    '    Range("C7:H2000").Select
    '    Selection.Sort Key1:=Range("C8"), Order1:=xlAscending, Header:=xlGuess, _
    '        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Columns("C:C").ColumnWidth = 25
    Columns("J:J").ColumnWidth = 8.5
    Range("h2").Select
    
End Sub




