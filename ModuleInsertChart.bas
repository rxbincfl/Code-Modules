Attribute VB_Name = "ModuleInsertChart"
'######################################################################################

Function ChangeChartAxisScale(CName As String, _
    Optional PlotAxis As Integer = xlPrimary, _
    Optional Xlower As Double = 0, _
    Optional Xupper As Double = 0, _
    Optional Ylower As Double = 0, _
    Optional YUpper As Double = 0, _
    Optional AxisColor As Double) As Variant

    Dim Rtn As String
    Dim MyAxis As Axis
    Set MyAxis = ActiveSheet.Shapes(CName).Chart.Axes(xlValue, PlotAxis)

    On Error Resume Next
    With MyAxis
        .Format.Line.ForeColor.RGB = AxisColor
        If Xlower = 0 Then
            .MinimumScaleIsAuto = True
            Rtn = "Xmin = auto"
        Else:
            .MinimumScale = Xlower
            Rtn = "Xmin = " & Xlower
        End If

        If Xupper = 0 Or (Xupper < Xlower) Then
            .MaximumScaleIsAuto = True
            Rtn = Rtn & "; Xmax = auto"
        Else
            .MaximumScale = Xupper
            Rtn = Rtn & "; Xmax = " & Xupper
        End If
    End With

    With MyAxis
        If Ylower = 0 Then
            .MinimumScaleIsAuto = True
            Rtn = Rtn & "; Ymin = auto"
        Else
            .MinimumScale = Ylower
            Rtn = Rtn & "; Ymin = " & Ylower
        End If

        If YUpper = 0 Or (Xupper < Xlower) Then
            .MaximumScaleIsAuto = True
            Rtn = Rtn & "; Ymax = auto"
        Else
            .MaximumScale = YUpper
            Rtn = Rtn & "; Ymax = " & YUpper
        End If
    End With

    ChangeChartAxisScale = Rtn
End Function
'    ch.Chart.Legend.Position = xlLegendPositionBottom

'######################################################################################
Public Sub InsertChart(rng As Range)
    'change log:
    '140905:    adding autorange to y-axis
    '######################################################################################
    Dim i As Long, lastrow As Long
    Dim rngSource As Range, rng1 As Range
    Dim ch As ChartObject
    Dim sc As SeriesCollection, srs As Series

    ' find the last row in the date column
    ' gotta find another dynamic way with Tables?
    lastrow = ActiveSheet.Cells(rows.Count, "a").End(xlUp).Row - 1

    ' add overlay chart
    Set rng1 = Range("a4:m23")
    Set ch = ActiveSheet.ChartObjects.Add _
    (Left:=rng1.Left, _
    Width:=rng1.Width, _
    Top:=rng1.Top, _
    Height:=rng1.Height)
    ch.Name = ActiveSheet.Name

    ' define the series collection
    Set sc = ch.Chart.SeriesCollection

    ' Header data
    Set rng1 = Range("2:2") ' row 2

    ' Add the Average Data Range (ADR) series
    i = WorksheetFunction.Match("ADR", rng1, 0)
    Set rngSource = Range( _
    Cells(rng1.Row + 1, i), _
    Cells(lastrow, i))
    rngSource.Select

    ' average of data range
    Set srs = sc.NewSeries
    With srs
        .Values = rngSource
        .Type = xlLine
        .MarkerStyle = xlMarkerStyleNone
        .Name = "ADR"
    End With

    With ch.Chart
        'chart name
        .HasTitle = True
        .ChartTitle.Characters.text = UCase(ActiveSheet.Name)

        'X axis name
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .AxisTitle.Characters.text = "Time - Days"
            .HasMajorGridlines = False
            .MajorTickMark = xlTickMarkNone
            .TickLabelPosition = xlNone
        End With

        'y-axis name
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Characters.text = UCase(srs.Name)
            .HasMajorGridlines = False
            .HasDisplayUnitLabel = False
            .MajorTickMark = xlTickMarkNone
            
        End With
    End With

    'ChangeChartAxisScale(CName As String, Optional Xlower As Double = 0, Optional Xupper As Double = 0, _
    Optional Ylower As Double = 0, Optional YUpper As Double = 0) As Variant
    Dim s As String
    s = ChangeChartAxisScale(ActiveSheet.Name, xlPrimary, , , _
        WorksheetFunction.Min(srs.Values), _
        WorksheetFunction.Max(srs.Values), vbBlue)

    ' Add the Close Data Range (Close) series
    ' add data to secondary y-axis
    i = WorksheetFunction.Match("Close", rng1, 0)
    Set rngSource = Range( _
    Cells(3, i), _
    Cells(lastrow, i))

    ' add trendline to "Close"
    Set srs = sc.NewSeries
    With srs
        .Values = rngSource
        .Type = xlLine
        .MarkerStyle = xlMarkerStyleNone
        .Name = rng1(1, i)
        .AxisGroup = xlSecondary
        .Trendlines.Add _
        Type:=xlLinear, _
        DisplayEquation:=True, _
        DisplayRSquared:=False
    End With

    s = ChangeChartAxisScale(ActiveSheet.Name, xlSecondary, , , _
        WorksheetFunction.Min(srs.Values), _
        WorksheetFunction.Max(srs.Values), vbRed)

    Dim t As Trendline
    Dim d As Double
    Dim r As Range

    Set t = sc("Close").Trendlines(1)
    ' make sure equation is displayed
    t.DisplayRSquared = True
    t.DisplayEquation = True
    t.DataLabel.Top = 0
    t.DataLabel.Left = 0

    ' select the trendline and try to move it
    'Application.ScreenUpdating = True
    Range("txt2col").Value = (t.DataLabel.text)
    
    ' move legend to bottom
    ch.Chart.Legend.Position = xlLegendPositionBottom
    'Application.ScreenUpdating = False

    Set ch = Nothing
    Set sc = Nothing
    Set srs = Nothing
    Set t = Nothing

End Sub
'######################################################################################

'######################################################################################
Sub EmbeddedChartFromScratch()
    Dim myChtObj As ChartObject
    Dim rngChtData As Range
    Dim rngChtXVal As Range
    Dim iColumn As Long

    ' make sure a range is selected
    If TypeName(Selection) <> "Range" Then Exit Sub

    ' define chart data
    Set rngChtData = Selection

    ' define chart's X values
    With rngChtData
        Set rngChtXVal = .Columns(1).Offset(1).Resize(.rows.Count - 1)
    End With

    ' add the chart
    Set myChtObj = ActiveSheet.ChartObjects.Add _
    (Left:=250, Width:=375, Top:=75, Height:=225)
    With myChtObj.Chart

        ' make an XY chart
        .ChartType = xlXYScatterLines

        ' remove extra series
        Do Until .SeriesCollection.Count = 0
            .SeriesCollection(1).Delete
        Loop

        ' add series from selected range, column by column
        For iColumn = 2 To rngChtData.Columns.Count
            With .SeriesCollection.NewSeries
                .Values = rngChtXVal.Offset(, iColumn - 1)
                .XValues = rngChtXVal
                .Name = rngChtData(1, iColumn)
            End With
        Next

    End With

End Sub


Public Sub SetYAxis()
    Dim MyAxis As Axis
    Set MyAxis = ActiveSheet.ChartObjects(1).Chart.Axes(xlValue, xlPrimary)
    With MyAxis    'Set properties of y-axis
        .HasMajorGridlines = True
        .HasTitle = True
        .AxisTitle.text = "My Y-Axis"
        .AxisTitle.Font.Color = vbBlack
        .TickLabels.Font.Color = vbBlue
        .MaximumScale = 200
    End With
End Sub

Public Sub SetXAxis()
    Dim MyAxis As Axis
    Set MyAxis = ActiveSheet.ChartObjects(1).Chart.Axes(xlValue, xlPrimary)
    With MyAxis    'Set properties of x-axis
        .HasMajorGridlines = False
        .HasTitle = True
        .AxisTitle.text = "My Axis"
        .AxisTitle.Font.Color = vbBlack
        '        .CategoryNames = Range("C2:C11")
        .TickLabels.Font.Color = vbBlue
    End With
End Sub

Sub InsertChart_()
    Dim Cht As Object
    Set Cht = ActiveSheet.ChartObjects.Add(Left:=3, Width:=300, Top:=10, Height:=300)
    With Cht
        '.Chart.SetSourceData Source:=ActiveSheet.Range("A3:C8")
        '.Chart.SeriesCollection(1).Type = xlColumn
        '.Chart.SeriesCollection(1).AxisGroup = 1
        '.Chart.SeriesCollection(2).Type = xlLine
        '.Chart.SeriesCollection(2).AxisGroup = 2
    End With
End Sub
