Attribute VB_Name = "CCHS_template"
Sub Invoice_Template(DateInputStart As Date, DateInputEnd As Date, CCHSAttendanceReport As String, CountID As Integer)
'Sub Invoice_Template()

    Dim myWSName As String
    Dim AppLog As String
    Dim myInvoice As String
    Dim LR As Long
    Dim LRNewSheet As Long
    Dim UpperCase As String
    Dim CCHSInvoiceReportMonthly As String

    'General functions
    monthly = Format(DateInputStart, "MMMM")
    
    Sheets(monthly & " Attendance Summary").Select
    CountID = CountID + 2
     'CCHSAttendanceReport = Sheets("main").Range("inputAttendanceTemplate").Value
'Associate
    
    'Filter by Associate only
    Rows("11:11").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=5, Criteria1:= _
        "Associate"
    'Filter criteria 18 only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:=">=18", Operator:=xlFilterValues

    'Select row and create new sheet
  
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
 
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("TotalData").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    LRNewSheet = Selection.Cells.Count
     If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("AssocTHC").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("AssocTHC").Select
        ActiveSheet.Paste
    End If
    
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    'Filter criteria 18 below only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:="<18", Operator:=xlFilterValues
    'Select row and create new sheet
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
 
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("TotalData").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    LRNewSheet = Selection.Cells.Count
     If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("AssocTHCLes").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("AssocTHCLes").Select
        ActiveSheet.Paste
    End If
    
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Selection.AutoFilter
    
'Senior Associate
    Rows("11:11").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=5, Criteria1:= _
        "Senior Associate"
    'Filter criteria 18 only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:=">=18", Operator:=xlFilterValues
    
    'Select row and create new sheet
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
 
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("TotalData").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    LRNewSheet = Selection.Cells.Count
     If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("SenAssocTHC").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("SenAssocTHC").Select
        ActiveSheet.Paste
    End If
       
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    'Filter criteria 18 below only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:="<18", Operator:=xlFilterValues
    'Select row and create new sheet
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
    
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("TotalData").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    LRNewSheet = Selection.Cells.Count
     If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("SenAssocTHCLess").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("SenAssocTHCLess").Select
        ActiveSheet.Paste
    End If
    
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Selection.AutoFilter
    
'Team Lead
    Rows("11:11").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=5, Criteria1:= _
        "Team Lead"
    'Filter criteria 18 only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:=">=18", Operator:=xlFilterValues

    'Select row and create new sheet
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
    
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("TotalData").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    LRNewSheet = Selection.Cells.Count
     If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("TeamLeadTHC").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("TeamLeadTHC").Select
        ActiveSheet.Paste
    End If
    
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

    
    'Filter criteria 18 below only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:="<18", Operator:=xlFilterValues
    'Select row and create new sheet
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
    
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("TotalData").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    LRNewSheet = Selection.Cells.Count
     If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("TeamLeadTHCLess").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("TeamLeadTHCLess").Select
        ActiveSheet.Paste
    End If
    
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Selection.AutoFilter

'Report Analyst
'Select row
    Rows("11:11").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=5, Criteria1:= _
        "Reports Analyst"
    'Filter criteria 18 only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:=">=18", Operator:=xlFilterValues
    
    'Select row and create new sheet
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
    
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("THC").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    LRNewSheet = Selection.Cells.Count
    If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
   
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("ReportAna").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("ReportAna").Select
        ActiveSheet.Paste
    End If
    
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    'Filter criteria 18 below only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:="<18", Operator:=xlFilterValues
     'Select row and create new sheet
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
    
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("TotalData").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
     LRNewSheet = Selection.Cells.Count
    If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
    
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("ReportAnaLess").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("ReportAnaLess").Select
        ActiveSheet.Paste
    End If
    
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Selection.AutoFilter
    
'QA Lead
    'Select row
    Rows("11:11").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=5, Criteria1:= _
        "QA Lead"
    'Filter criteria 18 only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:=">=18", Operator:=xlFilterValues
   
      'Select row and create new sheet
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
    
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("TotalData").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    LRNewSheet = Selection.Cells.Count
    If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
    
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("QALead").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("QALead").Select
        ActiveSheet.Paste
    End If
    
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

    
    'Filter criteria 18 below only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:="<18", Operator:=xlFilterValues
       'Select row and create new sheet
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
    
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("TotalData").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    LRNewSheet = Selection.Cells.Count
    If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
    
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("QALeadLess").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("QALeadLess").Select
        ActiveSheet.Paste
    End If
    
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Selection.AutoFilter

'Trainer Lead
    Rows("11:11").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=5, Criteria1:= _
        "Trainer Lead"
    'Filter criteria 18 only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:=">=18", Operator:=xlFilterValues
    
    'Select row and create new sheet
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
    
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("TotalData").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    LRNewSheet = Selection.Cells.Count
    If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
    
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("TrainerLead").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("TrainerLead").Select
        ActiveSheet.Paste
    End If
    
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    'Filter criteria 18 below only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:="<18", Operator:=xlFilterValues
   'Select row and create new sheet
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
    
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("TotalData").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    LRNewSheet = Selection.Cells.Count
    If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
    
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("TrainerLeadLess").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("TrainerLeadLess").Select
        ActiveSheet.Paste
    End If
    
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Selection.AutoFilter

'Team Supervisor
    Rows("11:11").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=5, Criteria1:= _
        "Team Supervisor"
    'Filter criteria 18 only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:=">=18", Operator:=xlFilterValues
      
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
    
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("TotalData").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    LRNewSheet = Selection.Cells.Count
    If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
    
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("TeamSuper").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("TeamSuper").Select
        ActiveSheet.Paste
    End If
    
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
 
    
    'Filter criteria 18 below only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:="<18", Operator:=xlFilterValues
        'Filter criteria 18 only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:=">=18", Operator:=xlFilterValues
      
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
    
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("TotalData").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    LRNewSheet = Selection.Cells.Count
    If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
    
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("TeamSuperLess").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("TeamSuperLess").Select
        ActiveSheet.Paste
    End If
    
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Selection.AutoFilter
    
'Operations Manager
    Rows("11:11").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=5, Criteria1:= _
        "Operations Manager"
    'Filter criteria 18 only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:=">=18", Operator:=xlFilterValues

    'Select row and create new sheet
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
    
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("TotalData").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    LRNewSheet = Selection.Cells.Count
    If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
    
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("OperationMang").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("OperationMang").Select
        ActiveSheet.Paste
    End If
    
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    'Filter criteria 18 below only
    ActiveSheet.Range("$A$11:$XFC$80").AutoFilter Field:=293, Criteria1:="<18", Operator:=xlFilterValues
    
   'Select row and create new sheet
    Range("E12:E" & CountID).Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste
    'Range("TotalData").Select
    
    Selection.Copy
    Range("THC").Select
    ActiveSheet.Paste
    If (Range("TotalData").Value <> "") Then
        Range("TotalData").Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    LRNewSheet = Selection.Cells.Count
    If (LRNewSheet > 5000) Then
        Range("THC").Select
        LRNewSheet = Selection.Cells.Count
    End If
    
    'count all data from a row
    If LRNewSheet = 2 Then
        Range("TCHT").Value = 0
        Range("TCHT").Select
        Selection.Copy
        Range("OperationMangLess").Select
        ActiveSheet.Paste
    Else
        Range("TCHT").Value = LRNewSheet
        Range("TCHT").Select
        Selection.Copy
        Range("OperationMangLess").Select
        ActiveSheet.Paste
    End If
    
    '''' remove data
    Range("TotalData:TCHT").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Selection.AutoFilter

'last part
    'create border
    Range("JW87:KD95").Select
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
    
    Range("JX86:KD95").Select
    Selection.Copy
    
    Sheets("Invoice_Template").Select
    Range("JX86:KD95").Select
    ActiveSheet.Paste
    
    'Create Total Headcount less/greater than
    Range("JX86").Select
    Range("JX86").Value = "Total HC (>18)"
    
    Range("JY95").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(R[-8]C,R[-7]C,R[-6]C,R[-5]C,R[-4]C,R[-3]C,R[-2]C,R[-1]C)"
    Range("JY96").Select
    
    Range("JX95").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(R[-8]C,R[-7]C,R[-6]C,R[-5]C,R[-4]C,R[-3]C,R[-2]C,R[-1]C)"
    
    Range("JX86:KD95").Select
    Selection.Copy
    
    Sheets("Invoice_Template").Select
    Range("JX86:KD95").Select
    ActiveSheet.Paste

    Sheets("Invoice").Select
    Range("D11") = Format(DateInputStart, "mmmm") & ". 1 to " & Format(DateInputStart, "mmmm") & ". " & Format(DateInputEnd, "dd") & ", " & Format(DateInputStart, "yyyy")
   
    
    'Unhide months
    If monthly = "June" Then
        'PO Usage sheet
        Sheets("PO Usage").Select
        Rows("8:14").Select
        Selection.EntireRow.Hidden = False
        'CE_50 AG HC_4501992666 sheet
        Sheets("CE_50 AG HC_4501992666").Select
        Rows("98:169").Select
        Selection.EntireRow.Hidden = False
        'Invoice sheet
        Sheets("Invoice").Select
    End If
      
    'Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
    Sheets("Invoice_Template").Select
    Range("JV87:KD95").Select
    Selection.Copy
    
    Sheets(monthly & " Attendance Summary").Select
    Range("JV87:KD95").Select
    ActiveSheet.Paste
    
    'Select Template
    Range("JV87:KD95").Select
    Selection.Copy

    Sheets("Invoice_Template").Select
    Range("JV87:KD95").Select
    ActiveSheet.Paste
    
    Sheets("Invoice").Select
   'Create formula and drag
        Range("L16").Select
    'ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-7],'Invoice_Template'!R86C282:R95C290,3,FALSE)"
    'senior
    ActiveCell.FormulaR1C1 = _
        "=IF(VLOOKUP(RC[-7],'Invoice_Template'!R86C282:R95C290,3,FALSE)>46,""46"",VLOOKUP(RC[-7],'Invoice_Template'!R86C282:R95C290,3,FALSE))"
    
    'Senior assoc
      Range("L17").Select
          ActiveCell.FormulaR1C1 = _
        "=IF(VLOOKUP(RC[-7],'Invoice_Template'!R86C282:R95C290,3,FALSE)>4,""4"",VLOOKUP(RC[-7],'Invoice_Template'!R86C282:R95C290,3,FALSE))"
    
    Range("L17:L23").Select
    Range("L23").Activate
    Selection.FillDown
    Range("I27").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-4],'Invoice_Template'!R86C282:R95C290,9,FALSE)"
    Range("I27:I29").Select
    Range("I29").Activate
    Selection.FillDown
    Range("K27").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-6],'Invoice_Template'!R86C282:R95C290,8,FALSE)"
    Range("K27:K29").Select
    Range("K29").Activate
    Selection.Cut
    Application.CutCopyMode = False
    Selection.FillDown
    Range("I31").Select
    ActiveCell.FormulaR1C1 = ""
    Range("I31").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-4],'Invoice_Template'!R86C282:R95C290,6,FALSE)"
    Range("I31:I32").Select
    Range("I32").Activate
    Selection.FillDown
    Range("J31").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-5],'Invoice_Template'!R86C282:R95C290,7,FALSE)"
    Range("J31:J32").Select
    Range("J32").Activate
    Selection.FillDown
    
    UpperCase = UCase(monthly)
    Range("C16:C23").Value = "HCS " & UpperCase & " SERVICE FEE"
    Range("C25").Value = "HCS " & UpperCase & " SERVICE FEE"
    Range("C27:C29").Value = "HCS " & UpperCase & " SERVICE FEE"

    Dim sourceSheet As Worksheet
    Dim dataSheet As Worksheet
    Dim nextRow As Integer
    
    'Set dataSheet = Sheets("PO Usage")
    Set sourceSheet = Sheets("CE_50 AG HC_4501992666")
    Set sourceSheet2 = Sheets("Invoice")
    Set dataSheet = Sheets("PO Usage")
    
    'Select sheet
    Sheets("PO Usage").Select
    'Search monthly service fee
    Cells.Find(What:="HCS " & monthly & " SERVICE FEE", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
    'Drag up
    Range(Selection, Selection.End(xlUp)).Select
    'Count drag
    nextRow = Selection.Cells.Count
    
    'Update month
    If monthly = "January" Then
        dataSheet.Cells(nextRow, 4).Value = sourceSheet.Range("O37").Value
         dataSheet.Cells(nextRow, 5).Value = sourceSheet2.Range("TotalSub").Value
    End If
     If monthly = "February" Then
        dataSheet.Cells(nextRow, 4).Value = sourceSheet.Range("O49").Value
         dataSheet.Cells(nextRow, 5).Value = sourceSheet2.Range("TotalSub").Value
    End If
     If monthly = "March" Then
        dataSheet.Cells(nextRow, 4).Value = sourceSheet.Range("O61").Value
        dataSheet.Cells(nextRow, 5).Value = sourceSheet2.Range("TotalSub").Value
    End If
     If monthly = "April" Then
        dataSheet.Cells(nextRow, 4).Value = sourceSheet.Range("O73").Value
         dataSheet.Cells(nextRow, 5).Value = sourceSheet2.Range("TotalSub").Value
    End If
     If monthly = "May" Then
        dataSheet.Cells(nextRow, 4).Value = sourceSheet.Range("O85").Value
         dataSheet.Cells(nextRow, 5).Value = sourceSheet2.Range("TotalSub").Value
    End If
     If monthly = "June" Then
        dataSheet.Cells(nextRow, 4).Value = sourceSheet.Range("O97").Value
         dataSheet.Cells(nextRow, 5).Value = sourceSheet2.Range("TotalSub").Value
    End If
    If monthly = "July" Then
        dataSheet.Cells(nextRow, 4).Value = sourceSheet.Range("O109").Value
         dataSheet.Cells(nextRow, 5).Value = sourceSheet2.Range("TotalSub").Value
    End If
     If monthly = "August" Then
        dataSheet.Cells(nextRow, 4).Value = sourceSheet.Range("O121").Value
         dataSheet.Cells(nextRow, 5).Value = sourceSheet2.Range("TotalSub").Value
    End If
     If monthly = "September" Then
        dataSheet.Cells(nextRow, 4).Value = sourceSheet.Range("O133").Value
         dataSheet.Cells(nextRow, 5).Value = sourceSheet2.Range("TotalSub").Value
    End If
     If monthly = "October" Then
        dataSheet.Cells(nextRow, 4).Value = sourceSheet.Range("O145").Value
         dataSheet.Cells(nextRow, 5).Value = sourceSheet2.Range("TotalSub").Value
    End If
     If monthly = "November" Then
        dataSheet.Cells(nextRow, 4).Value = sourceSheet.Range("O157").Value
         dataSheet.Cells(nextRow, 5).Value = sourceSheet2.Range("TotalSub").Value
    End If
     If monthly = "December" Then
        dataSheet.Cells(nextRow, 4).Value = sourceSheet.Range("O169").Value
         dataSheet.Cells(nextRow, 5).Value = sourceSheet2.Range("TotalSub").Value
    End If
    
    Sheets("Invoice").Select
    
    'select file
    monthly = Format(DateInputStart, "MMMM")
    yearly = Format(DateInputStart, "YYYY")
    CCHSInvoiceReportMonthly = "" & monthly & "\" & "Invoice_CCHS_PMFTC" & "_" & monthly & "_" & yearly & ""
    
    myInvoice = ActiveWorkbook.Name
    
    Sheets(monthly & " Attendance Summary").Select
    
    Range("JZ100:KA1000").Select
    Selection.Copy
    Range("TotalData").Select
    ActiveSheet.Paste

    Application.DisplayAlerts = False
    'ActiveWindow.Close savechanges:=True
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
   
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:= _
    "C:\CCHS Invoice Automation V2\output\" & CCHSInvoiceReportMonthly & " .xlsx", _
    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
     Application.DisplayAlerts = True
     
    'Create Applog
    Sheets(monthly & " Attendance Summary").Select
   
    Range("A9:B" & CountID + 3).Select
    Selection.Copy

    Workbooks.Add
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ChDir "C:\CCHS Invoice Automation V2\output\" & monthly
    ActiveWorkbook.SaveAs Filename:= _
        "C:\CCHS Invoice Automation V2\output\" & monthly & "\AppLog.xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
        
    Range("A2:B3").Select
    Selection.Delete Shift:=xlUp
    
    myWSName = ActiveWorkbook.Name
    
    Application.DisplayAlerts = False
    'ActiveWindow.Close savechanges:=True
    Workbooks(myWSName).Close SaveChanges:=True
    Application.DisplayAlerts = True

    AppLog = ActiveWorkbook.Name
    Application.DisplayAlerts = False
    'ActiveWindow.Close savechanges:=True
    Workbooks(AppLog).Close SaveChanges:=True
    Application.DisplayAlerts = True
    
    Application.DisplayAlerts = False
    Workbooks.Open "C:\CCHS Invoice Automation v2\Templates\Invoice_CCHS_PMFTC Inc.xlsm"
    Sheets(monthly & " Attendance Summary").Delete
    ActiveWorkbook.Close SaveChanges:=True
    Application.DisplayAlerts = True
   
End Sub


