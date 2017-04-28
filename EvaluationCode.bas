Attribute VB_Name = "EvaluationCode"


Sub Add_ListColumn_2_ExistingTable()
Dim oWS As Worksheet ' Worksheet Object
Dim oRange As Range ' Range Object - Contains Represents the List of Items that need to be made unique
Dim oLst As ListObject ' List Object
Dim oLC As ListColumn ' List Column Object
On Error GoTo Disp_Error
' ---------------------------------------------
' Coded by Shasur for www.vbadud.blogspot.com
' Modified by Rolando Brabant August 4, 2016
' ---------------------------------------------

    Set oWS = ActiveSheet
    If oWS.ListObjects.Count = 0 Then Exit Sub
    
    Set oLst = oWS.ListObjects(1)
    Set oLC = oLst.ListColumns.Add
    oLC.Name = "Total Price"
    oLC.DataBodyRange = "=[aapl]*[spy]"
    
    If Not oLC Is Nothing Then Set oLC = Nothing
    If Not oLst Is Nothing Then Set oLst = Nothing
    If Not oWS Is Nothing Then Set oWS = Nothing
    
    ' --------------------
    ' Error Handling
    ' --------------------
    
Disp_Error:
    If Err <> 0 Then
        MsgBox Err.Number & " - " & Err.Description, vbExclamation, "VBA Tips & Tricks Examples"
    Resume Next

End If
End Sub


'One of the most powerful commands in VBA: "EVALUATE" but hardly
 'anyone knows about it, understands it or uses it.
 
 'Can't use worksheet formulas directly in VBA right?
 'Run this macro:
 
Sub Neato()
    MsgBox Evaluate("SUM(A1:A10)")
End Sub
 
 
' 'Yeah I know, what about:
'Set Fn = Application.WorksheetFunction
'x = Fn.Sum(Range("A1:A10"))
'
' 'or if you prefer just:
'x = Application.Sum(Range("A1:A10"))
 
 '...but in most cases, why bother?
 
 
 'Another little known EVALUATE fact; you're familiar with the
 'shorthand brackets for referencing ranges right?
 
'Range("A1:A10").Select
'[A1:A10].Select
 
 'Did you know those brackets were shorthand for EVALUATE?
 
Sub NeatoNeato()
     
     'given...
    [A1:A10].Select
     
     'is the same as...
    Evaluate("A1:A10").Select
     
     'then this should work right?
    x = [SUM(A1:A10)]
    MsgBox x
     
     'or just...
    MsgBox [SUM(A1:A10)]
     
     'hey... you know with those brackets, it looks just like a cell
     'in VBA doesn't it? hehehehe...
     
End Sub

Sub MultiColumnTable_To_Array1()

Dim myTable As ListObject
Dim MyArray As Variant
Dim x As Long

'Set path for Table variable
  Set myTable = ActiveSheet.ListObjects("Table1")

'Create Array List from Table
  MyArray = myTable.DataBodyRange

'Loop through each item in Third Column of Table (displayed in Immediate Window [ctrl + g])
  For x = LBound(MyArray) To UBound(MyArray)
    Debug.Print MyArray(x, 3)
  Next x
  
End Sub

