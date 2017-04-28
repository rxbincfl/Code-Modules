Attribute VB_Name = "MyAverage"


Public Function MyAverage(rng As Range, val1 As Double)
Dim rng1 As Range, s1 As String
Dim rng2 As Range, s2 As String
Dim i As Long

    ' test data
    For i = 1 To rng.rows.Count
        rng.Cells(i) = WorksheetFunction.RandBetween(0, 100)
    Next i
    
    i = rng.rows.Count
    
    Set rng1 = rng.Range("a1:" & "a9")
    Set rng2 = rng.Range("a2:" & "a10")
    
    rng2.Cells.Value = rng1.Cells.Value
    
    ' now copy top value to stack
    rng.Cells(1, 1).Value = val1
    Range("c2").Value = WorksheetFunction.Sum(rng.Cells)
    
    'calculate average
    MyAverage = WorksheetFunction.Average(rng.Cells)
    Range("c3").Value = MyAverage
    
    s1 = rng1.Address
    s2 = rng.Address
        

End Function
