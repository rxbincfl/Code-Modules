Attribute VB_Name = "DownloadStockData"
'http://investexcel.net
Public MyString As String
Public myTable As ListObject
'July 4, 2016: Added break code for power query

Sub DownloadStockQuotes _
( _
    ByVal stockTicker As String, _
    ByVal StartDate As Date, _
    ByVal EndDate As Date, _
    ByVal DestinationCell As String, _
    ByVal freq As String _
)

    ' test of myaverage function
'    Dim i As Double
'    i = MyAverage(Range("a1:a10"), 54.12)
    
    Dim qurl As String
    Dim StartMonth, StartDay, StartYear, EndMonth, EndDay, EndYear As String
    StartMonth = Format(Month(StartDate) - 1, "00")
    StartDay = Format(Day(StartDate), "00")
    StartYear = Format(Year(StartDate), "00")

    EndMonth = Format(Month(EndDate) - 1, "00")
    EndDay = Format(Day(EndDate), "00")
    EndYear = Format(Year(EndDate), "00")
    qurl = "URL;http://real-chart.finance.yahoo.com/table.csv?s=" + stockTicker + "&a=" + StartMonth + "&b=" + StartDay + "&c=" + StartYear + "&d=" + EndMonth + "&e=" + EndDay + "&f=" + EndYear + "&g=" + freq + "&ignore=.csv"
    'http://real-chart.finance.yahoo.com/table.csv?s=RYL&a=00&b=11&c=2013&d=09&e=11&f=2014&g=d&ignore=.csv

    On Error GoTo ErrorHandler:

    With ActiveSheet.QueryTables.Add(Connection:=qurl, _
        destination:=Range(DestinationCell))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "20"
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With

ErrorHandler:

End Sub

Public Sub DownloadData()
    Dim frequency As String
    Dim NumRows As Integer
    Dim lastrow As Integer
    Dim stockTicker As String
    Dim i As Integer
    Dim ch As Chart

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set myTable = Sheets("parameters").ListObjects("tblSymbols")

    'frequency = Worksheets("Parameters").Range("b7")
    frequency = Range("Frequency").Value

    'Delete all sheets apart from Parameters and Collar sheets
    'Delete all chart sheets
    Dim ws As Worksheet
    Dim b1 As Boolean
    For Each ws In Worksheets
        b1 = (Not (ws.Name = "Parameters" _
        Or ws.Name = "ultimate oscillator" _
        Or ws.Name = "Version" _
        Or ws.Name = "FieldLookup" _
        Or ws.Name = "PQ"))
        If (b1) Then ws.Delete
    Next

    For Each ch In Charts
        ch.Delete
    Next ch

    'Application.DisplayAlerts = True

    'Loop through all tickers
    Dim oSh As Worksheet
    Set oSh = Worksheets("Parameters")
    oSh.Select

    updateStatusBar ("Downloading Data")
    'i = Range("firstSymbol").Row
    For Each Row In myTable.ListRows
        'select an entire column (data plus header)
        'code modified August 26, 2014 to work with tables
        stockTicker = Row.Range(1)

        If stockTicker = "" Then
            GoTo NextIteration
        End If

        Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = stockTicker

        Cells(1, 1) = "Stock Quotes for " & stockTicker

'=========================================================
        GoTo NextIteration ' exit early for power query
'=========================================================

        Call DownloadStockQuotes(stockTicker, _
        Range("Start_Date").Value, _
        Range("End_Date").Value, _
        "$a$2", frequency)
        
        Columns("a:a").TextToColumns destination:=Range("a1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, _
        FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), _
        Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1))

        Sheets(stockTicker).Columns("A:P").ColumnWidth = 10

            lastrow = Sheets(stockTicker).UsedRange.Row - 2 + Sheets(stockTicker).UsedRange.rows.Count
            If lastrow < 3 Then
                Application.DisplayAlerts = False
                Sheets(stockTicker).Delete
                GoTo NextIteration
                Application.DisplayAlerts = True
            End If

        ' add code to place the close of the day in the parameters sheet

        Row.Range(2).Value = Range("e3")

        lastrow = Range("a2").CurrentRegion.rows.Count
        Sheets(stockTicker).Sort.SortFields.Add Key:=Range("A3:A" & lastrow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        With Sheets(stockTicker).Sort
            .SetRange Range("A2:G" & lastrow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        Call MyADR(lastrow)

        '    MyTextToColumns
        MyTextToColumns (Row.Range(3))
        [txt2col].ClearContents

NextIteration:
        ' get the next symbol in the list
    Next Row

    updateStatusBar ("Done...")
    
'    GoTo ErrorHandler
'
'    If Sheets("Parameters").Range("d2") = True Then
'        On Error GoTo ErrorHandler:
'        Call CopyToCSV
'    End If
'
ErrorHandler:

    Worksheets("Parameters").Select
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub MultiColumnTable_To_Array()

    Dim myTable As ListObject
    Dim MyArray() As Variant
    Dim x As Long
    
    With Worksheets("PQ")
        .Activate
        
    End With
    
Exit Sub

    'Set path for Table variable

    'Create Array List from Table
    MyArray = myTable.DataBodyRange

    'Loop through each item in Third Column of Table (displayed in Immediate Window [ctrl + g])
    For x = LBound(MyArray) To UBound(MyArray)
        Debug.Print MyArray(x, 2)
    Next x

End Sub

Sub CopyToCSV()

    Dim MyPath As String
    Dim MyFileName As String

    dateFrom = Worksheets("Parameters").Range("start_date")
    dateTo = Worksheets("Parameters").Range("end_date")
    frequency = Worksheets("Parameters").Range("frequency")

    MyPath = "c:\temp\"

    For Each ws In Worksheets
        If ws.Name <> "Parameters" Then
            ticker = ws.Name
            MyFileName = ticker & " " & Format(dateFrom, "dd-mm-yyyy") & " - " & Format(dateTo, "dd-mm-yyyy") & " " & frequency
            If Not Right(MyPath, 1) = "\" Then MyPath = MyPath & "\"
            If Not Right(MyFileName, 4) = ".csv" Then MyFileName = MyFileName & ".csv"
            Sheets(ticker).Copy
            With ActiveWorkbook
                .SaveAs Filename:= _
                MyPath & MyFileName, _
                FileFormat:=xlCSV, _
                CreateBackup:=False
                .Close False
            End With
        End If
    Next

End Sub

Sub ResizeTable()

    Dim rng As Range
    Dim tbl As ListObject

    'Resize Table to 7 rows and 5 columns
    Set rng = Range("Table1[#All]").Resize(7, 5)

    ActiveSheet.ListObjects("Table1").Resize rng


    'Expand Table size by 10 rows
    Set tbl = ActiveSheet.ListObjects("Table1")

    Set rng = Range("Table1[#All]").Resize(tbl.Range.rows.Count + 10, tbl.Range.Columns.Count)

    tbl.Resize rng

End Sub
