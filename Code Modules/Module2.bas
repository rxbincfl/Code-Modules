Attribute VB_Name = "Module2"
Option Explicit
Dim CurrChart As Chart    'accessible to all procedures
Dim CurrSeries As Integer 'accessible to all procedures

Property Get Chart() As Chart
    Set Chart = CurrChart
End Property

Property Let Chart(Cht)
    Set CurrChart = Cht
End Property


Property Get ChartSeries()
    ChartSeries = CurrSeries
End Property

Property Let ChartSeries(SeriesNum)
    CurrSeries = SeriesNum
End Property


Property Get SeriesName() As Variant
    If SeriesNameType = "Range" Then
        Set SeriesName = Range(SERIESFormulaElement(CurrChart, CurrSeries, 1))
    Else
        SeriesName = SERIESFormulaElement(CurrChart, CurrSeries, 1)
    End If
End Property

Property Let SeriesName(SName)
    CurrChart.SeriesCollection(CurrSeries).Name = SName
End Property

Property Get SeriesNameType() As String
    SeriesNameType = SERIESFormulaElementType(CurrChart, CurrSeries, 1)
End Property


Property Get XValues() As Variant
    If XValuesType = "Range" Then
        Set XValues = Range(SERIESFormulaElement(CurrChart, CurrSeries, 2))
    Else
        XValues = SERIESFormulaElement(CurrChart, CurrSeries, 2)
    End If
End Property

Property Let XValues(XVals)
    CurrChart.SeriesCollection(CurrSeries).XValues = XVals
End Property

Property Get XValuesType() As String
    XValuesType = SERIESFormulaElementType(CurrChart, CurrSeries, 2)
End Property


Property Get Values() As Variant
    If ValuesType = "Range" Then
        Set Values = Range(SERIESFormulaElement(CurrChart, CurrSeries, 3))
    Else
        Values = SERIESFormulaElement(CurrChart, CurrSeries, 3)
    End If
End Property

Property Let Values(Vals)
    CurrChart.SeriesCollection(CurrSeries).Values = Vals
End Property

Property Get ValuesType() As String
    ValuesType = SERIESFormulaElementType(CurrChart, CurrSeries, 3)
End Property



Property Get PlotOrder()
    PlotOrder = SERIESFormulaElement(CurrChart, CurrSeries, 4)
End Property

Property Let PlotOrder(PltOrder)
    CurrChart.SeriesCollection(CurrSeries).PlotOrder = PltOrder
End Property

Property Get PlotOrderType() As String
    PlotOrderType = SERIESFormulaElementType(CurrChart, CurrSeries, 4)
End Property


Private Function SERIESFormulaElementType(ChartObj, SeriesNum, Element) As String
    '   Returns a string that describes the element of a chart's SERIES formula
    '   This function essentially parses and analyzes a SERIES formula

    '   Element 1: Series Name. Returns "Range" , "Empty", or "String"
    '   Element 2: XValues. Returns "Range", "Empty", or "Array"
    '   Element 3: Values. Returns "Range" or "Array"
    '   Element 4: PlotOrder. Always returns "Integer"

    Dim SeriesFormula As String
    Dim FirstComma As Integer, SecondComma As Integer, LastComma As Integer
    Dim FirstParen As Integer, SecondParen As Integer
    Dim FirstBracket As Integer, SecondBracket As Integer
    Dim StartY As Integer
    Dim SeriesName, XValues, Values, PlotOrder As Integer

    '   Exit if Surface chart (surface chrarts do not have SERIES formulas)
    If ChartObj.ChartType >= 83 And ChartObj.ChartType <= 86 Then
        SERIESFormulaElementType = "ERROR - SURFACE CHART"
        Exit Function
    End If

    '   Exit if nonexistent series is specified
    If SeriesNum > ChartObj.SeriesCollection.Count Or SeriesNum < 1 Then
        SERIESFormulaElementType = "ERROR - BAD SERIES NUMBER"
        Exit Function
    End If

    '   Exit if element is > 4
    If Element > 4 Or Element < 1 Then
        SERIESFormulaElementType = "ERROR - BAD ELEMENT NUMBER"
        Exit Function
    End If

    '   Get the SERIES formula
    SeriesFormula = ChartObj.SeriesCollection(SeriesNum).Formula

    '   Get the First Element (Series Name)
    FirstParen = InStr(1, SeriesFormula, "(")
    FirstComma = InStr(1, SeriesFormula, ",")
    SeriesName = Mid(SeriesFormula, FirstParen + 1, FirstComma - FirstParen - 1)
    If Element = 1 Then
        If IsRange(SeriesName) Then
            SERIESFormulaElementType = "Range"
        Else
            If SeriesName = "" Then
                SERIESFormulaElementType = "Empty"
            Else
                If TypeName(SeriesName) = "String" Then
                    SERIESFormulaElementType = "String"
                End If
            End If
        End If
        Exit Function
    End If

    '   Get the Second Element (X Range)
    If Mid(SeriesFormula, FirstComma + 1, 1) = "(" Then
        '       Multiple ranges
        FirstParen = FirstComma + 2
        SecondParen = InStr(FirstParen, SeriesFormula, ")")
        XValues = Mid(SeriesFormula, FirstParen, SecondParen - FirstParen)
        StartY = SecondParen + 1
    Else
        If Mid(SeriesFormula, FirstComma + 1, 1) = "{" Then
            '           Literal Array
            FirstBracket = FirstComma + 1
            SecondBracket = InStr(FirstBracket, SeriesFormula, "}")
            XValues = Mid(SeriesFormula, FirstBracket, SecondBracket - FirstBracket + 1)
            StartY = SecondBracket + 1
        Else
            '          A single range
            SecondComma = InStr(FirstComma + 1, SeriesFormula, ",")
            XValues = Mid(SeriesFormula, FirstComma + 1, SecondComma - FirstComma - 1)
            StartY = SecondComma
        End If
    End If
    If Element = 2 Then
        If IsRange(XValues) Then
            SERIESFormulaElementType = "Range"
        Else
            If XValues = "" Then
                SERIESFormulaElementType = "Empty"
            Else
                SERIESFormulaElementType = "Array"
            End If
        End If
        Exit Function
    End If

    '   Get the Third Element (Y Range)
    If Mid(SeriesFormula, StartY + 1, 1) = "(" Then
        '       Multiple ranges
        FirstParen = StartY + 1
        SecondParen = InStr(FirstParen, SeriesFormula, ")")
        Values = Mid(SeriesFormula, FirstParen + 1, SecondParen - FirstParen - 1)
        LastComma = SecondParen + 1
    Else
        If Mid(SeriesFormula, StartY + 1, 1) = "{" Then
            '           Literal Array
            FirstBracket = StartY + 1
            SecondBracket = InStr(FirstBracket, SeriesFormula, "}")
            Values = Mid(SeriesFormula, FirstBracket, SecondBracket - FirstBracket + 1)
            LastComma = SecondBracket + 1
        Else
            '          A single range
            FirstComma = StartY
            SecondComma = InStr(FirstComma + 1, SeriesFormula, ",")
            Values = Mid(SeriesFormula, FirstComma + 1, SecondComma - FirstComma - 1)
            LastComma = SecondComma
        End If
    End If
    If Element = 3 Then
        If IsRange(Values) Then
            SERIESFormulaElementType = "Range"
        Else
            SERIESFormulaElementType = "Array"
        End If
        Exit Function
    End If

    '   Get the Fourth Element (Plot Order)
    PlotOrder = Mid(SeriesFormula, LastComma + 1, Len(SeriesFormula) - LastComma - 1)
    If Element = 4 Then
        SERIESFormulaElementType = "Integer"
        Exit Function
    End If
End Function


Private Function SERIESFormulaElement(ChartObj, SeriesNum, Element) As String
    '   Returns one of four elements in a chart's SERIES formula (as a string)
    '   This function essentially parses and analyzes a SERIES formula

    '   Element 1: Series Name. Can be a range reference, a literal value, or nothing
    '   Element 2: XValues. Can be a range reference (including a non-contiguous range), a literal array, or nothing
    '   Element 3: Values. Can be a range reference (including a non-contiguous range), or a literal array
    '   Element 4: PlotOrder. Must be an integer

    Dim SeriesFormula As String
    Dim FirstComma As Integer, SecondComma As Integer, LastComma As Integer
    Dim FirstParen As Integer, SecondParen As Integer
    Dim FirstBracket As Integer, SecondBracket As Integer
    Dim StartY As Integer
    Dim SeriesName, XValues, Values, PlotOrder As Integer

    '   Exit if Surface chart (surface chrarts do not have SERIES formulas)
    If ChartObj.ChartType >= 83 And ChartObj.ChartType <= 86 Then
        SERIESFormulaElement = "ERROR - SURFACE CHART"
        Exit Function
    End If

    '   Exit if nonexistent series is specified
    If SeriesNum > ChartObj.SeriesCollection.Count Or SeriesNum < 1 Then
        SERIESFormulaElement = "ERROR - BAD SERIES NUMBER"
        Exit Function
    End If

    '   Exit if element is > 4
    If Element > 4 Then
        SERIESFormulaElement = "ERROR - BAD ELEMENT NUMBER"
        Exit Function
    End If

    '   Get the SERIES formula
    SeriesFormula = ChartObj.SeriesCollection(SeriesNum).Formula

    '   Get the First Element (Series Name)
    FirstParen = InStr(1, SeriesFormula, "(")
    FirstComma = InStr(1, SeriesFormula, ",")
    SeriesName = Mid(SeriesFormula, FirstParen + 1, FirstComma - FirstParen - 1)
    If Element = 1 Then
        SERIESFormulaElement = SeriesName
        Exit Function
    End If

    '   Get the Second Element (X Range)
    If Mid(SeriesFormula, FirstComma + 1, 1) = "(" Then
        '       Multiple ranges
        FirstParen = FirstComma + 2
        SecondParen = InStr(FirstParen, SeriesFormula, ")")
        XValues = Mid(SeriesFormula, FirstParen, SecondParen - FirstParen)
        StartY = SecondParen + 1
    Else
        If Mid(SeriesFormula, FirstComma + 1, 1) = "{" Then
            '           Literal Array
            FirstBracket = FirstComma + 1
            SecondBracket = InStr(FirstBracket, SeriesFormula, "}")
            XValues = Mid(SeriesFormula, FirstBracket, SecondBracket - FirstBracket + 1)
            StartY = SecondBracket + 1
        Else
            '          A single range
            SecondComma = InStr(FirstComma + 1, SeriesFormula, ",")
            XValues = Mid(SeriesFormula, FirstComma + 1, SecondComma - FirstComma - 1)
            StartY = SecondComma
        End If
    End If
    If Element = 2 Then
        SERIESFormulaElement = XValues
        Exit Function
    End If

    '   Get the Third Element (Y Range)
    If Mid(SeriesFormula, StartY + 1, 1) = "(" Then
        '       Multiple ranges
        FirstParen = StartY + 1
        SecondParen = InStr(FirstParen, SeriesFormula, ")")
        Values = Mid(SeriesFormula, FirstParen + 1, SecondParen - FirstParen - 1)
        LastComma = SecondParen + 1
    Else
        If Mid(SeriesFormula, StartY + 1, 1) = "{" Then
            '           Literal Array
            FirstBracket = StartY + 1
            SecondBracket = InStr(FirstBracket, SeriesFormula, "}")
            Values = Mid(SeriesFormula, FirstBracket, SecondBracket - FirstBracket + 1)
            LastComma = SecondBracket + 1
        Else
            '          A single range
            FirstComma = StartY
            SecondComma = InStr(FirstComma + 1, SeriesFormula, ",")
            Values = Mid(SeriesFormula, FirstComma + 1, SecondComma - FirstComma - 1)
            LastComma = SecondComma
        End If
    End If
    If Element = 3 Then
        SERIESFormulaElement = Values
        Exit Function
    End If

    '   Get the Fourth Element (Plot Order)
    PlotOrder = Mid(SeriesFormula, LastComma + 1, Len(SeriesFormula) - LastComma - 1)
    If Element = 4 Then
        SERIESFormulaElement = PlotOrder
        Exit Function
    End If
End Function

Private Function IsRange(ref) As Boolean
    '   Returns True if ref is a Range
    Dim x As Range
    On Error Resume Next
    Set x = Range(ref)
    If Err = 0 Then IsRange = True Else IsRange = False
End Function

