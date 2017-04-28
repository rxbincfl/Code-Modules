VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatrixMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class shows you how to calculate with matrices.
'You can sum up and multiply two  matrices. (MatrixA
'and MatrixB)To set the data call for Matrix A 'SetArrayA'
'with a one dimensional Array that will be the Rows,
'and the Column in which you want to have
'the Data. For Matrix B itïs the same (Call SetArrayB)

Option Explicit

Private MatrixA() As Double
Private MatrixB() As Double     ' Our two Matrices


Public Property Get MatrixARows() As Integer
' Gives you an Integer Value for the count of rows of
'Matrix AThe first row is 0 because itïs
'much easier to calculate 0-based.
    MatrixARows = UBound(MatrixA, 1)
End Property

Public Property Get MatrixBRows() As Integer
' Gives you an Integer value for the count of rows of
'Matrix B The first row is 0
    MatrixBRows = UBound(MatrixB, 1)
End Property

Public Property Get MatrixACols() As Integer
' Gives you an Integer Value for the count of columns of
'Matrix A The first column is 0
    MatrixACols = UBound(MatrixA, 2)
End Property

Public Property Get MatrixBCols() As Integer
' Gives you an Integer Value for the count of columns of
'Matrix B The first column is 0
    MatrixBCols = UBound(MatrixB, 2)

End Property

Public Function SetArrayA(InputArray() As Double, Col As Double)
' Set the Array for the Matrix A. InputArray() is onedimensional
' for the rows. If you want more than one row, so call it more
'than once.

' You have also to tell it in which column this rows should be.
    On Error Resume Next
    Dim i As Integer
    If UBound(MatrixA, 1) < 1 Then
        ReDim MatrixA(UBound(InputArray()), Col)
    Else
        ReDim Preserve MatrixA(UBound(InputArray()), Col)
    End If
    For i = 0 To UBound(InputArray())
        MatrixA(i, Col) = InputArray(i)
    Next i
    
End Function

Public Function SetArrayB(InputArray() As Double, Col As Integer)
' Same than SetArrayA but for the Matrix B
    On Error Resume Next
    Dim i As Integer
    If UBound(MatrixB, 1) < 1 Then
        ReDim MatrixB(UBound(InputArray()), Col)
    Else
        ReDim Preserve MatrixB(UBound(InputArray()), Col)
    End If
    For i = 0 To UBound(InputArray())
        MatrixB(i, Col) = InputArray(i)
    Next i
End Function

Public Function GetArrayA(OutputArray() As Double, Col As _
Integer)
' Gives you an one dimensional Array from a column.
    On Error Resume Next
    Dim i As Integer
    ReDim OutputArray(UBound(MatrixA, 1))
    For i = 0 To UBound(MatrixA, 1)
        OutputArray(i) = MatrixA(i, Col)
    Next i

End Function

Public Function GetArrayB(OutputArray() As Double, Col As _
Integer)
' Same
    On Error Resume Next
    Dim i As Integer
    ReDim OutputArray(UBound(MatrixB, 1))
    For i = 0 To UBound(MatrixB, 1)
        OutputArray(i) = MatrixB(i, Col)
    Next i
End Function

Function Addition(Result() As Double) As Boolean
' Sums up Matrix A and Matrix B
' Result is a twodimensional Array with the result.
' For Example:  Result(0,0) = 1         1   3
'               Result(1,0) = 2         2   4
'               Result(0,1) = 3
'               Result(1,1) = 4     will be the Result matrix
' The first index in the Result Array will be the row, and the
'second the column.

On Error GoTo errhandler

    Dim Row1() As Double, Row2() As Double, tmpRow1() As Double
    Dim tmpRow2() As Double
    Dim i As Integer
    Dim j As Integer
    
    If (MatrixARows <> MatrixBRows) Or _
       (MatrixACols <> MatrixBCols) Then GoTo errhandler

' NOTE: You can only summerize matrices when they are from the
'same dimension: A(2,3)+B(2,3) will work fine but
' A(4,5)+B(1,2) ist not possible.
    
    ReDim Row1(MatrixARows, MatrixACols)
    ReDim Row2(MatrixBRows, MatrixBCols)
    For i = 0 To MatrixACols
        GetArrayA tmpRow1(), i
        For j = 0 To UBound(tmpRow1())
            Row1(j, i) = tmpRow1(j)
        Next j
    Next i
    For i = 0 To MatrixBCols
        GetArrayB tmpRow2(), i
        For j = 0 To UBound(tmpRow2())
            Row2(j, i) = tmpRow2(j)
        Next j
    Next i
    
    ReDim Result(MatrixARows, MatrixACols)
    For i = 0 To MatrixARows
        For j = 0 To MatrixACols
            Result(i, j) = Row1(i, j) + Row2(i, j)
        Next j
    Next i
    Addition = True
    Exit Function
errhandler:
    Addition = False

End Function

Function Subtraction(Result() As Double) As Boolean

'
On Error GoTo errhandler

    Dim Row1() As Double, Row2() As Double, tmpRow1() As Double
    Dim tmpRow2() As Double
    Dim i As Integer
    Dim j As Integer
    
    If (MatrixARows <> MatrixBRows) Or _
       (MatrixACols <> MatrixBCols) Then GoTo errhandler
    
    ReDim Row1(MatrixARows, MatrixACols)
    ReDim Row2(MatrixBRows, MatrixBCols)
    For i = 0 To MatrixACols
        GetArrayA tmpRow1(), i
        For j = 0 To UBound(tmpRow1())
            Row1(j, i) = tmpRow1(j)
        Next j
    Next i
    For i = 0 To MatrixBCols
        GetArrayB tmpRow2(), i
        For j = 0 To UBound(tmpRow2())
            Row2(j, i) = tmpRow2(j)
        Next j
    Next i
    
    ReDim Result(MatrixARows, MatrixACols)
    For i = 0 To MatrixARows
        For j = 0 To MatrixACols
            Result(i, j) = Row1(i, j) - Row2(i, j)
        Next j
    Next i
    Subtraction = True
    Exit Function
errhandler:
    Subtraction = False

End Function

Function Multiplication(Result() As Double) As Boolean
' This is the function which multiplies two matrices.
' The Result Array looks like the Result Array from the Addition.

On Error GoTo errhandler

    Dim Row1() As Double, Row2() As Double, tmpRow1() As Double
    Dim tmpRow2() As Double
    Dim x As Integer
    Dim y As Integer
    Dim z As Integer
    Dim i As Integer, j As Integer, K As Integer
    If (MatrixACols <> MatrixBRows) Then GoTo errhandler

' NOTE: You can only multiple Matrices if there are as many columns in
' the first Matrix than rows in the second.

    i = MatrixARows: j = MatrixBCols: K = MatrixACols
    
    ReDim Row1(i, K) ' first Matrix
    ReDim Row2(K, j) ' second Matrix
    For x = 0 To K
        GetArrayA tmpRow1(), x
        For y = 0 To UBound(tmpRow1())
            Row1(y, x) = tmpRow1(y)
        Next y
    Next x
    For x = 0 To j
        GetArrayB tmpRow2(), x
        For y = 0 To UBound(tmpRow2())
            Row2(y, x) = tmpRow2(y)
        Next y
    Next x
    
    ReDim Result(i, j)
    Dim Sum As Double
    Sum = 0
    For x = 0 To i
        For y = 0 To j
            For z = 0 To K
                Sum = Sum + (Row1(x, z) * Row2(z, y))
            Next z
            Result(x, y) = Sum
            Sum = 0
        Next y
    Next x
        
    ' Everything fine
    Multiplication = True
    Exit Function
    
errhandler:
    ' Error
    Multiplication = False
End Function

Private Sub ClearUp()
    ReDim MatrixA(0, 0)
    ReDim MatrixB(0, 0)
End Sub

Private Sub Class_Initialize()
    ClearUp
End Sub

