VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScheduleSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub formatHeader()
    Worksheets("Schedule").Range("A2") = "Customer"
    Worksheets("Schedule").Range("B2") = "Part #"
    Worksheets("Schedule").Range("C2") = "Qty"
    Worksheets("Schedule").Range("D2") = "Required Time (hours)"
    Worksheets("Schedule").Range("E2") = "Machine"
End Sub
