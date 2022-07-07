VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Schedule_Days 
   Caption         =   "Make Schedule"
   ClientHeight    =   3684
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4776
   OleObjectBlob   =   "Schedule_Days.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Schedule_Days"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clearButton_Click()
    Unload Me
    Schedule_Days.Show
End Sub
Sub submitButtom_Click()
    Application.ScreenUpdating = False
    Dim Last_Row As Long
    Last_Row = Worksheets("ScheduleInfo").Cells(Rows.count, 1).End(xlUp).Row

    'clearing previous schedule
    Worksheets("Schedule").Range("F3", "FQ" & Last_Row + 1).Clear
    Call ScheduleScripts.markedCompleted
    Call ScheduleInfoSheet.sort5
    Call indivMachPriority
    Call ScheduleScripts.markedCompleted
    Call ScheduleInfoSheet.scheduleRowLabels
    Call ScheduleSheet.formatHeader
    Call UserForm
    Application.ScreenUpdating = True
End Sub
Sub UserForm()
    Dim days As String, machine As String
    If scheduleAll.Value = True Then
        machine = "gantry,sl-20,tl-2,tm-2,vf-2,vf-3,vf-4"
    Else
        machine = LCase(machineInput.Value)
    End If
    days = ""
    If Monday.Value = True Then
        days = days & "monday, "
    End If
    If Tuesday.Value = True Then
        days = days & "tuesday, "
    End If
    If Wednesday.Value = True Then
        days = days & "wednesday, "
    End If
    If Thursday.Value = True Then
        days = days & "thursday, "
    End If
    If Friday.Value = True Then
        days = days & "friday, "
    End If
    If Saturday.Value = True Then
        days = days & "saturday, "
    End If
    If Sunday.Value = True Then
        days = days & "sunday"
    End If
    Schedule_Days.Hide
    Call scheduleDayStorage(days, machine)
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = 1
        Schedule_Days.Hide
    End If
End Sub
Private Sub indivMachPriority()
    Dim Last_Row As Long, machine As String, machineArr As Variant, machineFirstRow As Long, machineLastRow As Long, machineName As Variant, firstCompletedRow As Long
    machine = "gantry,sl-20,tl-2,tm-2,vf-2,vf-3,vf-4"
    machineArr = Split(machine, ",")
    Last_Row = Worksheets("ScheduleInfo").Cells(Rows.count, 1).End(xlUp).Row
    Worksheets("ScheduleInfo").Range("A1:" & "G" & Last_Row).Borders.LineStyle = xlNone
    Worksheets("Schedule").Range("A1:E1048576").Borders.LineStyle = xlNone
    On Error Resume Next
    firstCompletedRow = Sheets("ScheduleInfo").Range("G:G").Find(what:="COMPLETED").Row
    If Err <> 0 Then
        firstCompletedRow = 1048576
        Err.Clear
    End If
    On Error GoTo 0
    For Each machineName In machineArr
        machineLastRow = Sheets("ScheduleInfo").Range("E1:E" & firstCompletedRow - 1).Find(what:=machineName, searchdirection:=xlPrevious).Row
        machineFirstRow = Sheets("ScheduleInfo").Range("E:E").Find(what:=machineName).Row
        Worksheets("ScheduleInfo").Range("F" & machineFirstRow).Value = 1
        Range("F" & machineFirstRow).Select
        Selection.AutoFill Destination:=Range("F" & machineFirstRow & ":" & "F" & machineLastRow), Type:=xlFillSeries
        With Worksheets("ScheduleInfo").Range("A" & machineFirstRow & ":" & "G" & machineLastRow)
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
        End With
        
        
    Next machineName
End Sub
