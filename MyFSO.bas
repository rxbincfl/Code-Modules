Attribute VB_Name = "MyFSO"
Option Explicit

Public fso As FileSystemObject

Sub AddOlEObject()

    Dim mainWorkBook As Workbook
    Dim Folderpath As String
    Dim NoOfFiles As Long
    Dim listfiles As Collection
    Dim fls As file
    Dim strCompFilePath As String
    Dim counter As Long

    Set mainWorkBook = ActiveWorkbook
    Sheets("Objects").Activate
    Folderpath = "C:\Users\Rolando\Pictures"
    Set fso = CreateObject("Scripting.FileSystemObject")
    NoOfFiles = fso.GetFolder(Folderpath).Files.Count
    Set listfiles = fso.GetFolder(Folderpath).Files
    For Each fls In listfiles
       strCompFilePath = Folderpath & "\" & Trim(fls.Name)
        If strCompFilePath <> "" Then
            If (InStr(1, strCompFilePath, "jpg", vbTextCompare) > 1 _
            Or InStr(1, strCompFilePath, "jpeg", vbTextCompare) > 1 _
            Or InStr(1, strCompFilePath, "png", vbTextCompare) > 1) Then
                 counter = counter + 1
                  Sheets("Objects").Range("A" & counter).Value = fls.Name
                  Sheets("Objects").Range("B" & counter).ColumnWidth = 25
                Sheets("Objects").Range("B" & counter).RowHeight = 100
                Sheets("Objects").Range("B" & counter).Activate
                Call insert(strCompFilePath, counter)
                Sheets("Objects").Activate
            End If
        End If
    Next
mainWorkBook.Save
End Sub

Function insert(PicPath, counter)
'MsgBox PicPath
    With ActiveSheet.Pictures.insert(PicPath)
        With .ShapeRange
            .LockAspectRatio = msoTrue
            .Width = 50
            .Height = 70
        End With
        .Left = ActiveSheet.Range("B" & counter).Left
        .Top = ActiveSheet.Range("B" & counter).Top
        .Placement = 1
        .PrintObject = True
    End With
End Function

