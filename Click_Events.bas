Attribute VB_Name = "Click_Events"
Option Explicit

Public Sub UpdateSymbols_Click()
    #If developMode = 1 Then
        Debug.Print "UpdateSymbols"
    #End If
    UpdateSymbols
End Sub


Public Sub GetNewData_Click()
    #If developMode = 1 Then
        Debug.Print "GetNewData"
    #End If
    Call GetData
End Sub
