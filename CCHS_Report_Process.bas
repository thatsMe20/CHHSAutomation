Attribute VB_Name = "CCHS_Report_Process"
Sub CCHS_Report(DateInputStart As Date, DateInputEnd As Date, CountID As Integer)

    'Sub CCHS_Report()
    Dim GetBook, CurrentSheet, wbSource, wbMain As String 'Declare local variables
    Dim GetYear As String
    Dim monthly As String
    Dim CCHSAttendanceReport As String
    Dim CCHSInvoiceReport As String
    Dim CountInvoiceTemp As Integer
    Dim ForCounts As Integer
    
    monthly = Format(DateInputStart, "MMMM")
    yearly = Format(DateInputStart, "YYYY")
    
    Windows("PMI CCHS Invoice Report Automation.xlsm").Activate
    
    'select file
    CCHSAttendanceReport = "C:\CCHS Invoice Automation V2\output\" & monthly & "\CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx"
    'CCHSAttendanceReport = Sheets("main").Range("inputAttendanceTemplate").Value
    CCHSInvoiceReport = Sheets("main").Range("inputInvoiceTemplate").Value
    
    'Open CCH Invoice Report Report
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=CCHSInvoiceReport
    Application.DisplayAlerts = True
   
    'Create month
    'Sheets.Add After:=ActiveSheet
    Sheets("Invoice_Template").Select
    Sheets("Invoice_Template").Copy Before:=Sheets(5)
    
    Sheets("Invoice_Template (2)").Select
    Sheets("Invoice_Template (2)").Name = monthly & " Attendance Summary"
    Sheets(monthly & " Attendance Summary").Select
    'ActiveWorkbook.Save
  
    
   'Open CCH Attendance Report
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    
    'Range("A9:MH120").Select
     'ActiveWorkbook.Close SaveChanges:=False
    Range("A9:N120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("A9:N120").Select
    ActiveSheet.Paste
    '2
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("Q12:Y120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("O12:W120").Select
    ActiveSheet.Paste
    '3
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("AB12:AJ120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("X12:AH1120").Select
    ActiveSheet.Paste
    '4
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("AM12:AU120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("AG12:AO120").Select
    ActiveSheet.Paste
    '5
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("AX12:BF120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("AP12:AX120").Select
    ActiveSheet.Paste
    '6
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("BI12:BQ120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("AY12:BG120").Select
    ActiveSheet.Paste
    '7
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("BT12:CB120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("BH12:BP120").Select
    ActiveSheet.Paste
    '8
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("CE12:CM120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("BQ12:BY120").Select
    ActiveSheet.Paste
    '9
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("CP12:CX120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("BZ12:CH120").Select
    ActiveSheet.Paste
    '10
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("DA12:DI120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("CI12:CQ120").Select
    ActiveSheet.Paste
    '11
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("DL12:DT120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("CR12:CZ120").Select
    ActiveSheet.Paste
    '12
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("DW12:DE120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("DA12:DI120").Select
    ActiveSheet.Paste
    '13
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("EH12:EP120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("DJ12:DR120").Select
    ActiveSheet.Paste
    '14
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("ES12:FA120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("DS12:EA120").Select
    ActiveSheet.Paste
    '15
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("FD12:FL120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("EB12:EJ120").Select
    ActiveSheet.Paste
    '16
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("FO12:FW120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("EK12:ES120").Select
    ActiveSheet.Paste
    '17
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("FZ12:GH120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("ET12:FB120").Select
    ActiveSheet.Paste
    '18
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("GK12:GS120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("FC12:FK120").Select
    ActiveSheet.Paste
    '19
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("GV12:HD120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("FL12:FT120").Select
    ActiveSheet.Paste
    '20
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("HG12:HO120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("FU12:GC120").Select
    ActiveSheet.Paste
    '21
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("HR12:HZ120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("GD12:GL120").Select
    ActiveSheet.Paste
    '22
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("IC12:IK120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("GM12:GU120").Select
    ActiveSheet.Paste
    '23
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("IN12:IV120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("GV12:HD120").Select
    ActiveSheet.Paste
    '24
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("IY12:JG120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("HE12:HM120").Select
    ActiveSheet.Paste
    '25
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("JJ12:JR120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("HN12:HV120").Select
    ActiveSheet.Paste
    '26
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("JU12:KC120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("HW12:IE120").Select
    ActiveSheet.Paste
    '27
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("KF12:KN120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("IF12:IN120").Select
    ActiveSheet.Paste
    '28
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("KQ12:KY120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("IO12:IW120").Select
    ActiveSheet.Paste
    '29
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("LB12:LJ120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("IX12:JF120").Select
    ActiveSheet.Paste
    '30
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("LM12:LU120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("JG12:JO120").Select
    ActiveSheet.Paste
    '31
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("LX12:MF120").Select
    Selection.Copy
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    Range("JP12:JX120").Select
    ActiveSheet.Paste
    

    'Hide unnecessary days
    If Range("IQ10").Value = DateInputEnd Then
        Columns("IX:JX").Select
        Selection.EntireColumn.Hidden = True
    End If
    
    If Range("JI10").Value = DateInputEnd Then
        Columns("JP:JX").Select
        Selection.EntireColumn.Hidden = True
    End If
    
   'Open CCH Attendance Report
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Range("MJ12:MN83").Select
    Selection.Copy
    
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    'Open CCH Invoice Report Report
    Sheets(monthly & " Attendance Summary").Select
    Range("JZ12").Select
    ActiveSheet.Paste
    
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Sheets(monthly).Select
    Range("MO12:MO83").Select
    Selection.Copy

    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    'Open CCH Invoice Report Report
    Sheets(monthly & " Attendance Summary").Select
    Range("KG12").Select
    ActiveSheet.Paste
    
    'Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    'Sheets(monthly).Select
    'Range("MP12:MP83").Select
    'Selection.Copy
    'Close Attendance Report
    'Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate
    'Open CCH Invoice Report Report
    'Sheets(monthly & " Attendance Summary").Select
    'Range("KH12").Select
    'ActiveSheet.Paste
    
    'Open CCH Attendance Report
    Windows("CCHS Attendance Report_" & monthly & "_" & yearly & ".xlsx").Activate
    Sheets(monthly).Select
    Range("PJ12:PN20").Select
    Selection.Copy
    ActiveWorkbook.Close SaveChanges:=False
    
    'Close Attendance Report
    Windows("Invoice_CCHS_PMFTC Inc.xlsm").Activate

    'Open CCH Invoice Report
     Sheets(monthly & " Attendance Summary").Select
    
    Range("TotalRegHrs").Select
    ActiveSheet.Paste
    
    'Open Invoice
    Sheets("Invoice").Select

    Call Invoice_Template(DateInputStart, DateInputEnd, CCHSAttendanceReport, CountID)
    
End Sub




