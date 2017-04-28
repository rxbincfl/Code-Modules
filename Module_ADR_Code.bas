Attribute VB_Name = "Module_ADR_Code"
'######################################################################################
' Modified by Rolando Brabant May 2, 2013
' Modified by Rolando Brabant July 8, 2015 : Added Hyperlink to Symbols
'######################################################################################
Public Sub MyADR(lastrow As Integer)
    Dim firstRow As Integer, iCount As Integer, MyAverage As Double
    Dim lo As ListObject
    

    ' Add the daily data range DR(i)=High(i)-Low(i)
    firstRow = 3
    lastrow = [a1].CurrentRegion.rows.Count
    iCount = Range("Frequency_Sample")
    Range("H2").Value = "DR"
    For i = firstRow To lastrow
        Range("H" & i) = Range("C" & i) - Range("D" & i)
    Next i

    'now add in the average adr
    Range("I2").Value = "ADR"
    For i = iCount + firstRow To lastrow
        MyAverage = WorksheetFunction.Average(Range("$H" & i - iCount, "$H" & i))
        Range("I" & i) = Format(MyAverage, "0.000")
    Next i

    'now calculate mean and std deviation of the ADR
    i = iCount + firstRow
    MyAverage = WorksheetFunction.Average(Range("$I" & i, "$I" & lastrow))
    Range("D1") = "ADR Mean"
    Range("E1") = Format(MyAverage, "0.000")
    Range("G1") = "ADR stdDev"
    MyAverage = WorksheetFunction.StDev(Range("$I" & i, "$I" & lastrow))
    Range("H1") = Format(MyAverage, "0.000")

    Dim ProftTarget As Double, StopLoss As Double
    Dim Slope As String

    Range("J1") = "Profit Target"
    ProftTarget = Range("$E1") * 0.15
    Range("K1") = Format(ProftTarget, "0.000")

    Range("L1") = "Stop Loss"
    StopLoss = Range("$E1") * 0.1
    Range("M1") = Format(StopLoss, "0.000")
    
    '   "I said it was ""awesome"" not ""awful""!!!"
    Range("O2").Formula = "=HYPERLINK(""[StockData.xlsm]Parameters!tblSymbols"")"

    '    Call InsertChart
    Call InsertChart(ActiveSheet.Range("K3"))
    '######################################################################################
End Sub
