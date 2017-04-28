Attribute VB_Name = "Statistics"

Public Sub CreateExcessReturnsMatrix()
Dim rng As Range, rngStart As Range
Dim Arr As Variant
Dim one As Variant
Dim ws As Worksheet


'this uses the Portfolio Power Query to load data
Set ws = Worksheets("Returns")
With ws 'w1
    .Activate
    .Cells.Clear
    
    Set rng = .[c2] 'reference cell
    
    'referenced to this sheet
    With Worksheets("PQ Data Pivot")
    ' this reads in the raw data via power query
        .PivotTables(1).TableRange1.Offset(2, 1).Copy rng(0, 0)
    End With
    
    Dim RawData As Range
    Set RawData = .UsedRange
    
    NumRows = RawData.rows.Count
    NumCols = RawData.Columns.Count
    
    RawData.Select 'make sure it is correct data
    
    ' this array is the returns for the period selected
    XArray = RawData.Value2
    
    
    
    ' two "" in a row inserts a "
    ' fill this range with 1/n
    'v = Evaluate("=IF(ISERROR(A1:K1), 13, 13)")
'    s1 = "=IF(ISERROR(RawData.address),1,1)"""
    
    'this is the Excess Returns data for calculating the Deviation matrix D
    Dim s1 As String, s2 As String
    s2 = "a1:a" & RawData.rows.Count
    s1 = "=IF(ISERROR(" & s2 & "),1,1)"
    
    ' this creates ones array
    ReDim one(1 To RawData.rows.Count, 1 To 1)
    Dim k As Variant
    ReDim k(1 To 1, 1 To 1) ' this is a constant!!
    k(1, 1) = 1 / NumRows
'    Debug.Print k(1, 1)
    
    one = Evaluate(s1)
    
    'set rng to col a
    Dim destination As Range
    NumRows = UBound(one, 1) - LBound(one, 1) + 1
    NumCols = UBound(one, 2) - LBound(one, 2) + 1
    
    Dim OnesArray As Variant
    ReDim OnesArray(1 To NumRows, 1 To NumCols)
    OnesArray = one
    
    'check on the ones array
    Set rng = Range("A1").Resize(NumRows, NumCols)
    rng.Select
    
    rng = OnesArray
    'equation for residual matrix: D = X-1*1tr*X
    Dim tmparray As Variant
    
'    Dim Fn As Object
    Set Fn = Application.WorksheetFunction
    
    '============test code===========
    Dim a, b, c 'all variants
    ReDim a(1 To 2, 1 To 3)
    a = [{3,5,1;9,1,4}]
    
    ReDim b(1 To 2, 1 To 1)
    b = [{1;1}]
    '============test code===========
    a = XArray
    b = OnesArray
    c = a
    
'    Cells.Clear
    
    'tmparray contains the averages
    a = Fn.MMult(Fn.MMult(b, Fn.Transpose(b)), a)
    
    NumRows = UBound(a, 1) - LBound(a, 1) + 1
    NumCols = UBound(a, 2) - LBound(a, 2) + 1
    
    Set rng = RawData.Resize(NumRows, NumCols)
    rng.Offset(NumRows + 1, 0) = a
    
    'now divide tmparray by num of rows
    'copy,select,pastespecial: divide
    
    Dim r1 As Range
    Dim r2 As Range
    rng(, UBound(a, 2) + 2).Value = NumRows
    Set r1 = rng.Offset(0, NumCols + 1).Resize(1, 1)
    
    'divide by numrows
    'this is rng now the average values
    Set rng = rng.Offset(NumRows + 1, 0)
    rng.Select
    r1.Copy
    rng.PasteSpecial _
        xlValues, xlPasteSpecialOperationDivide

    'subtract the averages matrix from the xmatrix
    'this rng is now the excess return values
    rng.Copy
    Set r1 = rng.Offset(NumRows + 1, 0)
    r1 = c
    r1.PasteSpecial xlValues, xlPasteSpecialOperationSubtract
    ReDim ExcessReturns(NumRows, NumCols)
    DArray = r1
    '============test code===========
'    Debug.Print TITLE
    '============test code===========
    

    'release memory
'===============================================================
'===============================================================
End With '\w1
End Sub

Public Sub CopyIndividualAssets()
With Worksheets("PQ")
Dim rng As Range
Dim e1 As Portfolios

    .Cells.Clear ' clear the WS
    .Activate ' the WS

    'copy/paste method
    With Worksheets("Pivot").PivotTables(1).TableRange1
        .Copy Range(StartAddress)
        
        
    End With
    Set rng = .UsedRange
    rng.Name = "MyPortfolioStats"
    rng.Select
    NumRows = rng.rows.Count
    NumCols = rng.Columns.Count + 1 'account for sharpe ratio col
   
    Dim Arr As Variant ' declare an unallocated dynamic array.
    Set rng = rng.Resize(NumRows, NumCols)
    
    Averages = rng.Resize(rng.rows.Count - 1, 1).Offset(1, 1)
    StDevs = rng.Resize(rng.rows.Count - 1, 1).Offset(1, 2) ' standard deviation array
    
    
    'now add sharpe ratio column

    'labels to help analysis
    'now copy the portfolio column to the constraint area
    'save the average column in a global variable for stats
    rng.Columns(2).Copy Range(StartAddress).Offset(4, NumCols + 1)
    rng.Columns(3).Copy Range(StartAddress).Offset(4, NumCols)
    Set rng = Range(StartAddress)
    
    With rng
        .Offset(-2, 1).Value = ASSETS
        .Offset(-2, NumCols + 3).Value = PORTFOLIO
        .Offset(-3, NumCols + 0).Value = TITLE
        .Offset(1, NumCols + 1).Value = "Constraining Variable"
        .Offset(2, NumCols + 1).Value = "Value of Constraint"
        
        .Offset(1, NumCols + 2).Value = "None"
        .Offset(2, NumCols + 2).Value = "N/A"
        .Offset(1, NumCols + 3).Value = "at sigma <="
        .Offset(2, NumCols + 3).Value = "2.838%"
        .Offset(1, NumCols + 4).Value = "at MU ="
        .Offset(2, NumCols + 4).Value = "1.462%"
        .Offset(1, NumCols + 5).Value = "None"
        .Offset(2, NumCols + 5).Value = "N/A"
            
        .Offset(0, NumCols - 1).Value = "mu/sigma"
        .Offset(NumRows + 4, NumCols + 1).Value = "Sigma Wi"
        .Offset(NumRows + 5, NumCols + 1).Value = "MU"
        .Offset(NumRows + 6, NumCols + 1).Value = "co-variance"
        .Offset(NumRows + 7, NumCols + 1).Value = "MU/Sigma"
        
        .Offset(0, NumCols + equalwt).Value = "Equal Wt."
        .Offset(0, NumCols + MaxRtn).Value = "Max Ret."
        .Offset(0, NumCols + MinStDev).Value = "Min St Dev"
        .Offset(0, NumCols + MaxSR).Value = "Max SR"
    End With
    
    ' fill in the equal wt with initial values in %
    Set rng = rng.Offset(5, NumCols + equalwt).Resize(NumRows - 1, 1)
    rng.Value = 1 / rng.rows.Count
    rng.NumberFormat = "0.00%"
End With
End Sub

