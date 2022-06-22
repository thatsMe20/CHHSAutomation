Attribute VB_Name = "CCHS_Attendance"
Sub Attendance(DateInputStart As Date, DateInputEnd As Date)
    
    
    Dim AttFilterMonth As String
    Dim InputMonth As String
    Dim DateInput As String
    Dim CCHSReportAutomation As String
    Dim MasterList As String
    Dim CCHSAttendanceReport As String
    Dim X As Integer
    Dim TotalCountDTR As Integer
    Dim PlusOne As Integer
    Dim ValidForRowSix As String
    Dim DTRPath As String
    Dim DTRPath2 As String
    Dim CountTotalA As Integer
    Dim monthly As String
    Dim CountID As Integer
    Dim CCHSAttendanceReportMonthly As String
    Dim yearly As String
    Dim newMonth As String
    
    'select file
    CCHSAttendanceSchedule = "C:\CCHS Invoice Automation V2\Templates\CCHS Attendance Schedule.xlsx"
    CCHSReportAutomation = "C:\CCHS Invoice Automation V2\Templates\CCHS Report Automation.xlsm"
    MasterList = "C:\CCHS Invoice Automation V2\Templates\Master List.xlsx"
    CCHSAttendanceReport = Sheets("main").Range("inputAttendanceTemplate").Value
    
    'select file
    monthly = Format(DateInputStart, "MMMM")
    yearly = Format(DateInputStart, "YYYY")
    CCHSAttendanceReportMonthly = "" & monthly & "\" & "CCHS Attendance Report" & "_" & monthly & "_" & yearly & ""
    
    'Open this to display active employee
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=MasterList
    Application.DisplayAlerts = True
    
    Sheets("HC Dump").Select
    
    'date format march-2022
    Range("W1").Value = Format(DateInputStart, "mm-dd-yyyy")
    Range("W1").Select
    Selection.NumberFormat = "[$-3409]dd-mmm;@"
    
    'Filter by month
    Range("E1").Select
    Selection.AutoFilter
    AttFilterMonth = Range("X1").Value
    Range("X1").Select
    ActiveSheet.Range("$A$1:$X$2892").AutoFilter Field:=5, Criteria1:= _
        AttFilterMonth
    
    'change date format to month only
    InputMonth = Format(AttFilterMonth, "mmmm")

    'copy data and paste to CCHS Report Automation
    'Name
    Range("F1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    'Paste currently in 10000 row
    Range("A10000").Select
    ActiveSheet.Paste
    
    'Employee ID
    Range("G1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    'Paste currently in 10000 row
    Range("B10000").Select
    ActiveSheet.Paste
    
    'Role Title
    Range("J1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    'Paste currently in 10000 row
    Range("C10000").Select
    ActiveSheet.Paste
    
    'Select 3 row
    Range("A10000").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    'Close masterlist
    Application.DisplayAlerts = False
    ActiveWindow.Close
    Application.DisplayAlerts = True
    
    Sheets("validation").Select
    
    'remove duplicate data
    Range("N1").Select
    ActiveSheet.Paste
    
    Range("J2:L300").Select
    Selection.ClearContents
    
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("pasteData").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("$J$1:$L$258").RemoveDuplicates Columns:=Array(1, 2, 3), _
        Header:=xlYes
    Range("L2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("E2").Select
    ActiveSheet.Paste
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.SmallScroll Down:=-3
    Range("A2").Select
    ActiveSheet.Paste
    
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B2").Select
    ActiveSheet.Paste
    
    
    'Open CCH Attendance Report
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=CCHSAttendanceSchedule
    Application.EnableEvents = True
    
    Sheets("Attendance Schedule").Select
    
    Range("A1").Select
    str1 = Left(DateInputEnd, 2)
    strTotal = str1 * 2
    Cells.Find(What:=monthly, After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
    , SearchFormat:=False).Activate
    Range("A2").Value = ActiveCell.Column
        
    'Input Column Number
    ColNumber = Range("A2").Value + 1
    'Convert To Column Letter
    ColLetter = Split(Cells(1, ColNumber).Address, "$")(1)
    
    ColNumber2 = ColNumber + strTotal - 1
    ColLastLetter = Split(Cells(1, ColNumber2).Address, "$")(1)
    
    'Total count of employee
    Range("A7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Counts = Selection.Cells.Count + 6
    
    'Select and Copy Data
    Range(ColLetter & 5, ColLastLetter & Counts).Select
    Selection.Copy
        
    'Open CCH Attendance Report
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=CCHSAttendanceReport
    Application.EnableEvents = True
        
    Sheets("Schedule").Select
     
        Range("DT5").Select
        ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
            
    Windows("CCHS Attendance Schedule.xlsx").Activate
    Range("A7:E" & Counts).Select
    Selection.Copy
    
    Application.DisplayAlerts = False
    ActiveWindow.Close
    Application.DisplayAlerts = True
    
    Windows("CCHS Attendance Report.xlsx").Activate
    Range("A7").Select
    ActiveSheet.Paste

    'Create month
    Sheets("Template").Select
    Sheets("Template").Copy After:=Sheets(3)
        
    Sheets("Template (2)").Select
    'On Error Resume Next
    Sheets("Template (2)").Name = monthly
    Sheets(monthly).Select
    
    Range("H10").Value = Format(DateInputStart, "mm-dd-yyyy")

    'Hide unnecessary days
     If Range("KS10").Value = DateInputEnd Then
         Columns("LB:LJ").Select
         Selection.EntireColumn.Hidden = True
         
         Columns("LM:LU").Select
         Selection.EntireColumn.Hidden = True
         
         Columns("LX:MF").Select
         Selection.EntireColumn.Hidden = True
     End If
     
     If Range("LO10").Value = DateInputEnd Then
         Columns("LX:MF").Select
         Selection.EntireColumn.Hidden = True
     End If
     
    Windows("PMI CCHS Invoice Report Automation.xlsm").Activate
    'copy then paste to attendance summary report
    Sheets("validation").Select
    Range("A2:E2").Select
    Range("E2").Activate
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Copy
    
    Windows("CCHS Attendance Report.xlsx").Activate
    Sheets(monthly).Select
    
    Range("A12:E12").Select
    ActiveSheet.Paste
    
    Range("A12").Select
    Range(Selection, Selection.End(xlDown)).Select
    CountID = Selection.Cells.Count + 11
    'Selection.End(xlDown).Select
    
    'Drag data from PR to PV
    Range("PR12:PV" & CountID).Select
    Selection.FillDown
    
    Range("MJ12:MO" & CountID).Select
    Selection.FillDown
    
    Range("MQ12:NV" & CountID).Select
    Selection.FillDown
    
    Range("F12:MH" & CountID).Select
    Selection.FillDown
        
     'Create border
    Range("A12:E" & CountID).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    'Close Attendance Report
    'ActiveWorkbook.SaveAs Filename:=CCHSAttendanceReportMonthly
    Sheets(monthly).Select
    Range("A12:MH" & CountID).Select
    Range("A12:MH" & CountID).Value = Range("A12:MH" & CountID).Value
    
    ActiveWorkbook.SaveAs Filename:= _
        "C:\CCHS Invoice Automation V2\output\" & CCHSAttendanceReportMonthly & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

   Call CCHS_Report(DateInputStart, DateInputEnd, CountID)
    
End Sub
