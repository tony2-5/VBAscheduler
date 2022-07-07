Attribute VB_Name = "ScheduleScripts"
Public timeSet As Boolean
Public cellRow As Long
Sub markedCompleted()
    Dim setPriority As Integer, Last_Row As Long
    Last_Row = Worksheets("ScheduleInfo").Cells(Rows.count, 1).End(xlUp).Row
    For i = 2 To Last_Row
        If Worksheets("ScheduleInfo").Range("G" & i).Value = "COMPLETED" Then
            Worksheets("ScheduleInfo").Range("F" & i).ClearContents
        End If
    Next i
End Sub
Sub scheduleDayStorage(days As String, machine As String)
    Dim machineArr As Variant
    Dim machineString As String
    machine = Replace(machine, ", ", ",")
    machine = LCase(machine)
    machineArr = Split(machine, ",")
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
    Dim storedDays As String
    storedDays = days
    For m = LBound(machineArr) To UBound(machineArr)
        cellRow = Worksheets("Schedule").Range("E:E").Find(what:=machineArr(m)).Row
        machineString = machineArr(m)
        Call scheduleSetUp(days, machineString)
    Next m
    
End Sub

Sub scheduleSetUp(days As String, machine As String)
    'initializing variables
    Dim Last_Row As Long, machineName As String, machineLastRow As Long, machineFirstRow As Long
    Dim daysArr As Variant, color As Integer
    Dim count As Long, count2 As Long
    'machines and their colors: gantry(red),sl-20(green),tl-2(yellow),tm-2(cyan),vf-2(purple),vf-3(gold),vf-4(orange)
    Select Case machine
        Case Is = "gantry"
            color = 3
        Case Is = "sl-20"
            color = 10
        Case Is = "tl-2"
            color = 6
        Case Is = "tm-2"
            color = 8
        Case Is = "vf-2"
            color = 13
        Case Is = "vf-3"
            color = 12
        Case Is = "vf-4"
            color = 55
    End Select
    
    'starting cell row in schedule
    
    Last_Row = Worksheets("ScheduleInfo").Cells(Rows.count, 1).End(xlUp).Row
    
    'getting machine name to create array with in our inventory sheet
    machineName = machine
    machineLastRow = Sheets("ScheduleInfo").Range("E:E").Find(what:=machineName, searchdirection:=xlPrevious).Row
    machineFirstRow = Sheets("ScheduleInfo").Range("E:E").Find(what:=machineName).Row
    'array to initialize partRowArr() size based on amount of machines with specified name
    count = 0
    For Z = machineFirstRow To machineLastRow
        If Worksheets("ScheduleInfo").Range("E" & Z) = machineName And Worksheets("ScheduleInfo").Range("G" & Z) = "IN QUEUE" Then
            count = count + 1
        End If
    Next Z
    
    Dim partRowArr() As Integer
    ReDim partRowArr(count)
    
    'Filling array with hour value from inventory sheet from highest on list to lowest
    count2 = 0
    For Z2 = machineFirstRow To machineLastRow
        If Worksheets("ScheduleInfo").Range("E" & Z2) = machineName And Worksheets("ScheduleInfo").Range("G" & Z2) = "IN QUEUE" Then
            partRowArr(count2) = Worksheets("ScheduleInfo").Range("D" & Z2).Value
            count2 = count2 + 1
        End If
    Next Z2
    
    'getting days from userform and making array days()
    daysArr = Split(days, ", ")
    Call scheduleBuild(partRowArr(), daysArr, color)
End Sub
Sub scheduleBuild(partRowArr() As Integer, days As Variant, color As Integer)
    Dim tempTime As Integer, leftOver As Integer, difference As Integer, count As Integer
    uB = UBound(days)
    lB = LBound(days)
    uB1 = UBound(partRowArr)
    count = 0
    
    'filling schedule for each day user specifies
    For i = lB To uB
        Select Case days(i)
        
        'MONDAY
        
        Case Is = "monday"
            Dim startTime As Integer, endTime As Integer, time As Integer
            If count = 0 And timeSet = True Then
                startTime = Format(Now, "H")
                count = count + 1
            Else
                startTime = 0
            End If
            endTime = 23
            time = (endTime - startTime) + 1
            tempTime = time
            Dim sumHolder As Integer, valueCheck As Integer
            
            
            valueCheck = 0
            sumHolder = 0

            For lb1 = LBound(partRowArr) To uB1
                valueCheck = valueCheck + partRowArr(lb1)
                If valueCheck > tempTime Then
                    valueCheck = valueCheck - partRowArr(lb1)
                    difference = tempTime - valueCheck
                    leftOver = partRowArr(lb1) - difference
                    partRowArr(lb1) = difference
                    valueCheck = valueCheck + partRowArr(lb1)
                    tempTime = 100
                End If
                If valueCheck <= time Then
                    If partRowArr(lb1) <> 0 Or tempTime = 100 Then
                        If partRowArr(lb1) <> 0 Then
                            For l1 = 0 To partRowArr(lb1) - 1
                                Worksheets("Schedule").Cells(cellRow, startTime + 6 + l1 + sumHolder).Interior.ColorIndex = color
                            Next l1
                            sumHolder = sumHolder + partRowArr(lb1)
                            partRowArr(lb1) = 0
                        End If
                        If tempTime = 100 Then
                            partRowArr(lb1) = leftOver
                            leftOver = 0
                            difference = 0
                            cellRow = cellRow - 1
                        End If
                        cellRow = cellRow + 1
                    End If
                End If
            Next lb1
            
        'TUESDAY
            
        Case Is = "tuesday"
            Dim startTime2 As Integer, endTime2 As Integer, time2 As Integer
            If count = 0 And timeSet = True Then
                startTime2 = Format(Now, "H")
                count = count + 1
            Else
                startTime2 = 0
            End If
            endTime2 = 23
            time2 = (endTime2 - startTime2) + 1
            tempTime = time2
            
            Dim sumHolder2 As Integer, valueCheck2 As Integer
            
            valueCheck2 = 0
            sumHolder2 = 0
            
            For lb1 = LBound(partRowArr) To uB1
                valueCheck2 = valueCheck2 + partRowArr(lb1)
                If valueCheck2 > tempTime Then
                    valueCheck2 = valueCheck2 - partRowArr(lb1)
                    difference = tempTime - valueCheck2
                    leftOver = partRowArr(lb1) - difference
                    partRowArr(lb1) = difference
                    valueCheck2 = valueCheck2 + partRowArr(lb1)
                    tempTime = 100
                End If
                If valueCheck2 <= time2 Then
                    If partRowArr(lb1) <> 0 Or tempTime = 100 Then
                        If partRowArr(lb1) <> 0 Then
                            For l2 = 0 To partRowArr(lb1) - 1
                                Worksheets("Schedule").Cells(cellRow, startTime2 + 30 + l2 + sumHolder2).Interior.ColorIndex = color
                            Next l2
                            sumHolder2 = sumHolder2 + partRowArr(lb1)
                            partRowArr(lb1) = 0
                        End If
                        If tempTime = 100 Then
                            partRowArr(lb1) = leftOver
                            leftOver = 0
                            difference = 0
                            cellRow = cellRow - 1
                        End If
                        cellRow = cellRow + 1
                    End If
                End If
            Next lb1
            
        'WEDNESDAY
            
        Case Is = "wednesday"
            Dim startTime3 As Integer, endTime3 As Integer, time3 As Integer
            If count = 0 And timeSet = True Then
                startTime3 = Format(Now, "H")
                count = count + 1
            Else
                startTime3 = 0
            End If
            endTime3 = 23
            time3 = (endTime3 - startTime3) + 1
            tempTime = time3
    
            Dim sumHolder3 As Integer, valueCheck3 As Integer
            
            valueCheck3 = 0
            sumHolder3 = 0

            For lb1 = LBound(partRowArr) To uB1
                valueCheck3 = valueCheck3 + partRowArr(lb1)
                If valueCheck3 > tempTime Then
                    valueCheck3 = valueCheck3 - partRowArr(lb1)
                    difference = tempTime - valueCheck3
                    leftOver = partRowArr(lb1) - difference
                    partRowArr(lb1) = difference
                    valueCheck3 = valueCheck3 + partRowArr(lb1)
                    tempTime = 100
                End If
                If valueCheck3 <= time3 Then
                    If partRowArr(lb1) <> 0 Or tempTime = 100 Then
                        If partRowArr(lb1) <> 0 Then
                            For l3 = 0 To partRowArr(lb1) - 1
                                Worksheets("Schedule").Cells(cellRow, startTime3 + 54 + l3 + sumHolder3).Interior.ColorIndex = color
                            Next l3
                        End If
                        sumHolder3 = sumHolder3 + partRowArr(lb1)
                        partRowArr(lb1) = 0
                        If tempTime = 100 Then
                            partRowArr(lb1) = leftOver
                            leftOver = 0
                            difference = 0
                            cellRow = cellRow - 1
                        End If
                        cellRow = cellRow + 1
                    End If
                End If
            Next lb1
            
        'THURSDAY
            
        Case Is = "thursday"
            Dim startTime4 As Integer, endTime4 As Integer, time4 As Integer
            If count = 0 And timeSet = True Then
                startTime4 = Format(Now, "H")
                count = count + 1
            Else
                startTime4 = 0
            End If
            endTime4 = 23
            time4 = (endTime4 - startTime4) + 1
            tempTime = time4
    
            Dim sumHolder4 As Integer, valueCheck4 As Integer
            
            valueCheck4 = 0
            sumHolder4 = 0

            For lb1 = LBound(partRowArr) To uB1
                valueCheck4 = valueCheck4 + partRowArr(lb1)
                If valueCheck4 > tempTime Then
                    valueCheck4 = valueCheck4 - partRowArr(lb1)
                    difference = tempTime - valueCheck4
                    leftOver = partRowArr(lb1) - difference
                    partRowArr(lb1) = difference
                    valueCheck4 = valueCheck4 + partRowArr(lb1)
                    tempTime = 100
                End If
                If valueCheck4 <= time4 Then
                    If partRowArr(lb1) <> 0 Or tempTime = 100 Then
                        If partRowArr(lb1) <> 0 Then
                            For l4 = 0 To partRowArr(lb1) - 1
                                Worksheets("Schedule").Cells(cellRow, startTime4 + 78 + l4 + sumHolder4).Interior.ColorIndex = color
                            Next l4
                            sumHolder4 = sumHolder4 + partRowArr(lb1)
                            partRowArr(lb1) = 0
                        End If
                        If tempTime = 100 Then
                            partRowArr(lb1) = leftOver
                            leftOver = 0
                            difference = 0
                            cellRow = cellRow - 1
                        End If
                        cellRow = cellRow + 1
                    End If
                End If
            Next lb1
            
            
        'FRIDAY
            
        Case Is = "friday"
          Dim startTime5 As Integer, endTime5 As Integer, time5 As Integer
          If count = 0 And timeSet = True Then
                startTime5 = Format(Now, "H")
                count = count + 1
            Else
                startTime5 = 0
            End If
            endTime5 = 23
            time5 = (endTime5 - startTime5) + 1
            tempTime = time5
    
            Dim sumHolder5 As Integer, valueCheck5 As Integer
            
            valueCheck5 = 0
            sumHolder5 = 0

            For lb1 = LBound(partRowArr) To uB1
                valueCheck5 = valueCheck5 + partRowArr(lb1)
                If valueCheck5 > tempTime Then
                    valueCheck5 = valueCheck5 - partRowArr(lb1)
                    difference = tempTime - valueCheck5
                    leftOver = partRowArr(lb1) - difference
                    partRowArr(lb1) = difference
                    valueCheck5 = valueCheck5 + partRowArr(lb1)
                    tempTime = 100
                End If
                If valueCheck5 <= time5 Then
                    If partRowArr(lb1) <> 0 Or tempTime = 100 Then
                        If partRowArr(lb1) <> 0 Then
                            For l5 = 0 To partRowArr(lb1) - 1
                                Worksheets("Schedule").Cells(cellRow, startTime5 + 102 + l5 + sumHolder5).Interior.ColorIndex = color
                            Next l5
                            sumHolder5 = sumHolder5 + partRowArr(lb1)
                            partRowArr(lb1) = 0
                        End If
                        If tempTime = 100 Then
                            partRowArr(lb1) = leftOver
                            leftOver = 0
                            difference = 0
                            cellRow = cellRow - 1
                        End If
                        cellRow = cellRow + 1
                    End If
                End If
            Next lb1
          
        'SATURDAY
        
        Case Is = "saturday"
            Dim startTime6 As Integer, endTime6 As Integer, time6 As Integer
            If count = 0 And timeSet = True Then
                startTime6 = Format(Now, "H")
                count = count + 1
            Else
                startTime6 = 0
            End If
            endTime6 = 23
            time6 = (endTime6 - startTime6) + 1
            tempTime = time6
    
            Dim sumHolder6 As Integer, valueCheck6 As Integer
            
            valueCheck6 = 0
            sumHolder6 = 0

            For lb1 = LBound(partRowArr) To uB1
                valueCheck6 = valueCheck6 + partRowArr(lb1)
                If valueCheck6 > tempTime Then
                    valueCheck6 = valueCheck6 - partRowArr(lb1)
                    difference = tempTime - valueCheck6
                    leftOver = partRowArr(lb1) - difference
                    partRowArr(lb1) = difference
                    valueCheck6 = valueCheck6 + partRowArr(lb1)
                    tempTime = 100
                End If
                If valueCheck6 <= time6 Then
                    If partRowArr(lb1) <> 0 Or tempTime = 100 Then
                        If partRowArr(lb1) <> 0 Then
                            For l6 = 0 To partRowArr(lb1) - 1
                                Worksheets("Schedule").Cells(cellRow, startTime6 + 126 + l6 + sumHolder6).Interior.ColorIndex = color
                            Next l6
                            sumHolder6 = sumHolder6 + partRowArr(lb1)
                            partRowArr(lb1) = 0
                        End If
                        If tempTime = 100 Then
                            partRowArr(lb1) = leftOver
                            leftOver = 0
                            difference = 0
                            cellRow = cellRow - 1
                        End If
                        cellRow = cellRow + 1
                    End If
                End If
            Next lb1
            
        'SUNDAY
        
        Case Is = "sunday"
            Dim startTime7 As Integer, endTime7 As Integer, time7 As Integer
            If count = 0 And timeSet = True Then
                startTime7 = Format(Now, "H")
                count = count + 1
            Else
                startTime7 = 0
            End If
            endTime7 = 23
            time7 = (endTime7 - startTime7) + 1
            tempTime = time7
    
            Dim sumHolder7 As Integer, valueCheck7 As Integer
            
            valueCheck7 = 0
            sumHolder7 = 0

            For lb1 = LBound(partRowArr) To uB1
                valueCheck7 = valueCheck7 + partRowArr(lb1)
                If valueCheck7 > tempTime Then
                    valueCheck7 = valueCheck7 - partRowArr(lb1)
                    difference = tempTime - valueCheck7
                    leftOver = partRowArr(lb1) - difference
                    partRowArr(lb1) = difference
                    valueCheck7 = valueCheck7 + partRowArr(lb1)
                    tempTime = 100
                End If
                If valueCheck7 <= time7 Then
                    If partRowArr(lb1) <> 0 Or tempTime = 100 Then
                        If partRowArr(lb1) <> 0 Then
                            For l7 = 0 To partRowArr(lb1) - 1
                                Worksheets("Schedule").Cells(cellRow, startTime7 + 150 + l7 + sumHolder7).Interior.ColorIndex = color
                            Next l7
                            sumHolder7 = sumHolder7 + partRowArr(lb1)
                            partRowArr(lb1) = 0
                        End If
                        If tempTime = 100 Then
                            partRowArr(lb1) = leftOver
                            leftOver = 0
                            difference = 0
                            cellRow = cellRow - 1
                        End If
                        cellRow = cellRow + 1
                    End If
                End If
            Next lb1
        End Select
    Next i
End Sub
Sub currentTimeTF(bool2 As Boolean)
    timeSet = bool2
    Call ScheduleInfoSheet.refreshButton_Click
End Sub
