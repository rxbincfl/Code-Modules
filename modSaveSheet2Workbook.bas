Attribute VB_Name = "modSaveSheet2Workbook"
Option Explicit
Const sLoc As String = "E:\Dropbox\investment data\data\Schwab\Trades"
#Const clrConns = 11


Public Sub cmdSaveSht2Wbk()
Dim sht As Worksheet
Dim wbk As Workbook
Dim wbkc As WorkbookConnection

    'clear all connections
    #If clrConns = 1 Then
       Call ClearConnections
    #End If
    
    

    'this routine will save worksheet to workbook for _
    easier maintenance of positions
    
    'copy each worksheet to new/existing workbook
    '
    For Each sht In ActiveWorkbook.Worksheets
        Debug.Print sht.Name
        
        If sht.[a1] = "Long Put" Then
            Debug.Print "yes"
            Call CreateWorkbook(wbkc)
        Else
            Debug.Print "no"
        End If
        
    Next
    
    
End Sub

Public Sub CreateWorkbook(conn As WorkbookConnection)
    'if the workbook does not exist then create it
    'otherwise, add(replace) the new pivot data to the workbook
Dim objWBConnect As WorkbookConnection

Set objWBConnect = ThisWorkbook.Connections.Add( _
    Name:="New Connection", _
    Description:="My New Connection Demo", _
    ConnectionString:="OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source=C:\Files\Northwind 2007.accdb", _
    CommandText:="SELECT [First Name], [Last Name] FROM Customers", _
    lCmdtype:=xlCmdSql)
    
    Debug.Print objWBConnect.Name
    
End Sub




Public Sub showConnections()
Dim conn As WorkbookConnection
For Each conn In ActiveWorkbook.Connections
  Debug.Print conn.Name
  Call ClearConnections
Next conn
End Sub


