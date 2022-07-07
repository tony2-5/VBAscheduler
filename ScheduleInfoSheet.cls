VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScheduleInfoSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub currentTime_Click()
    Call currentTimeTF(currentTime.Value)
End Sub
Private Sub priorityButton_Click()
    Dim topNum As String, Last_Row As Long
    Last_Row = Worksheets("ScheduleInfo").Cells(Rows.count, 1).End(xlUp).Row
    topNum = InputBox("Enter Part# To Set to Top", "Top Priority")
    If topNum <> "" Then
    For i = 2 To Last_Row
        If Worksheets("ScheduleInfo").Cells(i, 6).Value = 1 Then
            Worksheets("ScheduleInfo").Cells(i, 6).Value = "2"
        End If
    Next i
    
    For i2 = 2 To Last_Row
        If Worksheets("ScheduleInfo").Cells(i2, 2).Value = topNum Then
             Worksheets("ScheduleInfo").Cells(i2, 6).Value = "1"
        End If
    Next i2
    Call Schedule_Days.submitButtom_Click
    End If
End Sub
Sub sort5()
    Dim Last_Row As Long
    Last_Row = Worksheets("ScheduleInfo").Cells(Rows.count, 1).End(xlUp).Row
    Worksheets("ScheduleInfo").Range("A1:G" & Last_Row).Select
    ActiveWorkbook.Worksheets("ScheduleInfo").sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ScheduleInfo").sort.SortFields.Add2 Key:=Range("G2:G" & Last_Row) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "IN QUEUE,COMPLETED", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("ScheduleInfo").sort.SortFields.Add2 Key:=Range("E2:E" & Last_Row) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("ScheduleInfo").sort.SortFields.Add2 Key:=Range("F2:F" & Last_Row) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ScheduleInfo").sort
        .SetRange Range("A1:G" & Last_Row)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Private Sub scheduleButton_Click()
    Schedule_Days.Show
End Sub
Sub refreshButton_Click()
    Schedule_Days.submitButtom_Click
End Sub
Sub scheduleRowLabels()
    Dim machineArr As Variant, x As Long, y As Long, TempTxt1 As String, TempTxt2 As String, machineStart As Long, machineEnd As Long, count As Integer
    Dim firstMachineStart As Long, prevMachineEnd As Long, Last_Row As Long, hold As Long
    
    Last_Row = Worksheets("Schedule").Cells(Rows.count, 1).End(xlUp).Row
    
    Dim completeArr As Variant
    completeArr = Array("gantry", "sl-20", "tl-2", "tm-2", "vf-2", "vf-3", "vf-4")
    If Schedule_Days.scheduleAll.Value = True Then
         machine = "gantry,sl-20,tl-2,tm-2,vf-2,vf-3,vf-4"
    Else
        machine = Schedule_Days.machineInput.Value
    End If
    
    machine = Replace(machine, ", ", ",")
    machineArr = Split(machine, ",")

    'Alphabetize Sheet Names in machineArr
    
    For x = LBound(machineArr) To UBound(machineArr)
        For y = x To UBound(machineArr)
            If UCase(machineArr(y)) < UCase(machineArr(x)) Then
                TempTxt1 = machineArr(x)
                TempTxt2 = machineArr(y)
                machineArr(x) = TempTxt2
                machineArr(y) = TempTxt1
            End If
        Next y
    Next x
    count = 0
    For m = LBound(machineArr) To UBound(machineArr)
        On Error Resume Next
        count = WorksheetFunction.Match(machineArr(m), completeArr, 0)
        If Err <> 0 Then
                MsgBox ("Enter machine names as seen under 'Schedule Colors'. If entering multiple machines make sure to use commas. Ex: gantry,vf-2,sl-20")
                End
            End If
        On Error GoTo 0
        firstMachineStart = 3
        machineStart = (Range("E:E").Find(what:=completeArr(count - 1)).Row)
        If machineArr(m) = completeArr(UBound(completeArr)) Then
            On Error Resume Next
            machineEnd = Worksheets("ScheduleInfo").Range("G:G").Find(what:="COMPLETED").Row - 1
            If Err <> 0 Then
                machineEnd = Sheets("ScheduleInfo").Range("E:E").Find(what:=completeArr(count - 1), searchdirection:=xlPrevious).Row
                Err.Clear
            End If
            On Error GoTo 0
        Else
            machineEnd = (Range("E:E").Find(what:=completeArr(count)).Row) - 1
        End If
        If m = 0 Then
            Sheets("ScheduleInfo").Select
            Range("A" & machineStart & ":" & "A" & machineEnd).Select
            Selection.Copy
            Sheets("Schedule").Select
            Worksheets("Schedule").Range("A" & (firstMachineStart) & ":" & "A" & firstMachineStart + (machineEnd - machineStart)).PasteSpecial
            Sheets("ScheduleInfo").Select
            Worksheets("ScheduleInfo").Range("B" & machineStart & ":" & "B" & machineEnd).Copy
            Sheets("Schedule").Select
            Worksheets("Schedule").Range("B" & (firstMachineStart) & ":" & "B" & firstMachineStart + (machineEnd - machineStart)).PasteSpecial
            Sheets("ScheduleInfo").Select
            Worksheets("ScheduleInfo").Range("C" & machineStart & ":" & "C" & machineEnd).Copy
            Sheets("Schedule").Select
            Worksheets("Schedule").Range("C" & (firstMachineStart) & ":" & "C" & firstMachineStart + (machineEnd - machineStart)).PasteSpecial
            Sheets("ScheduleInfo").Select
            Worksheets("ScheduleInfo").Range("D" & machineStart & ":" & "D" & machineEnd).Copy
            Sheets("Schedule").Select
            Worksheets("Schedule").Range("D" & (firstMachineStart) & ":" & "D" & firstMachineStart + (machineEnd - machineStart)).PasteSpecial
            Sheets("ScheduleInfo").Select
            Worksheets("ScheduleInfo").Range("E" & machineStart & ":" & "E" & machineEnd).Copy
            Sheets("Schedule").Select
            Worksheets("Schedule").Range("E" & (firstMachineStart) & ":" & "E" & firstMachineStart + (machineEnd - machineStart)).PasteSpecial
            prevMachineEnd = firstMachineStart + (machineEnd - machineStart)
        Else
            Sheets("ScheduleInfo").Select
            Range("A" & machineStart & ":" & "A" & machineEnd).Select
            Selection.Copy
            Sheets("Schedule").Select
            Worksheets("Schedule").Range("A" & prevMachineEnd + 1 & ":" & "A" & prevMachineEnd + 1 + (machineEnd - machineStart)).PasteSpecial
            Sheets("ScheduleInfo").Select
            Worksheets("ScheduleInfo").Range("B" & machineStart & ":" & "B" & machineEnd).Copy
            Sheets("Schedule").Select
            Worksheets("Schedule").Range("B" & prevMachineEnd + 1 & ":" & "B" & prevMachineEnd + 1 + (machineEnd - machineStart)).PasteSpecial
            Sheets("ScheduleInfo").Select
            Worksheets("ScheduleInfo").Range("C" & machineStart & ":" & "C" & machineEnd).Copy
            Sheets("Schedule").Select
            Worksheets("Schedule").Range("C" & prevMachineEnd + 1 & ":" & "C" & prevMachineEnd + 1 + (machineEnd - machineStart)).PasteSpecial
            Sheets("ScheduleInfo").Select
            Worksheets("ScheduleInfo").Range("D" & machineStart & ":" & "D" & machineEnd).Copy
            Sheets("Schedule").Select
            Worksheets("Schedule").Range("D" & prevMachineEnd + 1 & ":" & "D" & prevMachineEnd + 1 + (machineEnd - machineStart)).PasteSpecial
            Sheets("ScheduleInfo").Select
            Worksheets("ScheduleInfo").Range("E" & machineStart & ":" & "E" & machineEnd).Copy
            Sheets("Schedule").Select
            Worksheets("Schedule").Range("E" & prevMachineEnd + 1 & ":" & "E" & prevMachineEnd + 1 + (machineEnd - machineStart)).PasteSpecial
            prevMachineEnd = prevMachineEnd + 1 + (machineEnd - machineStart)
        End If
        Next m
        Worksheets("Schedule").Range("A" & prevMachineEnd + 1 & ": E1048576").ClearContents
        Sheets("ScheduleInfo").Select
End Sub


