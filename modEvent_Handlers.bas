Attribute VB_Name = "modEvent_Handlers"
Option Explicit

Public RunWhen As Double
Public Const cRunIntervalSeconds = 120 ' two minutes
Public Const cRunWhat = "modGetData.GetData"  ' the name of the procedure to run

Dim sn() As Variant

Public Sub InitTimer(secs As Double, fn As String)
Static i As Long
    ReDim Preserve sn(i)
    sn(i) = Array(DateAdd("s", secs, Time))
    Application.OnTime sn(i), fn
    i = i + 1

'    sn = Array _
'    ( _
'        DateAdd("s", 1, Time), _
'        DateAdd("s", 3, Time), _
'        DateAdd("s", 6, Time) _
'    )
    
'    Application.OnTime sn(1), "M_snb_ontime_start_2"
'    Application.OnTime sn(2), "M_snb_ontime_start_3"
End Sub

Public Sub CancelTimer()
'    Application.OnTime sn(0), "M_snb_ontime_start_1", , False
'    Application.OnTime sn(1), "M_snb_ontime_start_2", , False
'    Application.OnTime sn(2), "M_snb_ontime_start_3", , False
End Sub

Sub M_snb_ontime_start_1()
    ThisWorkbook.Sheets("sheet1").TextBox1.Text = Time
    sn(0) = DateAdd("s", 1, Time)
    Application.OnTime sn(0), "M_snb_ontime_start_1"
End Sub

Sub M_snb_ontime_start_2()
    ThisWorkbook.Sheets("sheet1").TextBox2.Text = Time
    sn(1) = DateAdd("s", 3, Time)
    Application.OnTime sn(1), "M_snb_ontime_start_2"
End Sub

Sub M_snb_ontime_start_3()
    ThisWorkbook.Sheets("Sheet1").TextBox3.Text = Time
    sn(2) = DateAdd("s", 6, Time)
    Application.OnTime sn(2), "M_snb_ontime_start_3"
End Sub

'================================================================================
'================================================================================
'Starting A Timer
'To start a repeatable timer, create a procedure named StartTimer as shown below:
Sub StartTimer11()
    RunWhen = Now + TimeSerial(0, 0, cRunIntervalSeconds)
    Application.OnTime _
        EarliestTime:=RunWhen, _
        Procedure:=cRunWhat, _
        Schedule:=True
End Sub

'Stopping A Timer
'   At some point, you or your code will need to terminate the OnTime schedule loop. _
    To cancel a pending OnTime event, you must provide the exact time that it is scheduled to run. _
    That is the reason we stored the time in the Public variable RunWhen. _
    You can think of the RunWhen value as a unique key into the OnTime settings. _
    (It is certainly possible to have multiple OnTime events pending. _
    In this, you should store each procedure's scheduled time in a separate variable. _
    Each OnTime event needs its own RunWhen value.) _
    The code below illustrates how to stop a pending OnTime event.
    
Sub StopTimer11()
    On Error Resume Next
    Application.OnTime _
        EarliestTime:=RunWhen, _
        Procedure:=cRunWhat, _
        Schedule:=False
End Sub



