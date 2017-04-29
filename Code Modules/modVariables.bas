Attribute VB_Name = "modVariables"
Option Explicit

Public ws As Worksheet

Public loSet As ListObject
Public loGet As ListObject
Public loTest As ListObject
Public rCols As Range
Public rRows As Range
Public rBody As Range
Public rData As Range
Public rHeader As Range
Public rStart As Range

Public iCol As Long
Public iRow As Long
Public iStep As Long
Public iLastRow As Long
Public iRowCnt As Long
Public iColCnt As Long

Public sMsg As String

Public Const sSheetName As String = "Sheet1"
Public Const sTableName As String = "Table1"
Public Const sTableName2 As String = "Table2"
Public Const NL As String = vbNewLine
Public Const DNL As String = vbNewLine & vbNewLine

