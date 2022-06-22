Attribute VB_Name = "BROWSEFILES"
Sub SelectFile()

Dim DialogBox As FileDialog
Dim path As String

Set DialogBox = Application.FileDialog(msoFileDialogFilePicker)
DialogBox.Title = "Select file for " & FileType
DialogBox.Filters.Clear
DialogBox.Show

If DialogBox.SelectedItems.Count = 1 Then
path = DialogBox.SelectedItems(1)
End If

Sheets("Main").Range("InputDTRTemplate").Value = path
End Sub
Sub RunMyCodeNow()
Call Selectdates
End Sub
Sub Selectdates()
    Dim DateInputStart As Date
    Dim DateInputEnd As Date
    Dim Dateformat As String
    Dim Dateformat1 As String
    Dim Dateformatdmmyy As String
    Dim DateInputUpload As Date
    Dim DateFilterbyMonth As Date
    Dim dtrconcat As String
    Dim CountDTRTemp As String
    Dim DTRHiddenTemplate As String
    Dim DTRHTemplate1 As String
    Dim DTRHTemplate2 As String
    Dim MainDTRReport As String
    Dim CountA As Integer
    
    
    'Change the DTR name and path
    MainDTRReport = Sheets("main").Range("InputDTRTemplate").Value
    DateInputStart = Sheets("main").Range("DateStart").Value
    DateInputEnd = Sheets("main").Range("DateEnd").Value
    
    'date mmddyyyy
    'MainDTRReport
    Dateformat = Right(MainDTRReport, 15)
    Dateformat1 = Left(dtrconcat, 10)
    Dateformatdmmyy = Format(dtrconcat1, "dd/mm/yyyy")
    
    'select file
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=MainDTRReport
    Application.DisplayAlerts = True
    monthly = Format(DateInputStart, "MMMM")
     
    
    'save DTR
    dtrconcat = Right(MainDTRReport, 19)
    DTRReport = "" & monthly & "\" & dtrconcat
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:= _
    "C:\CCHS Invoice Automation V2\output\" & DTRReport, _
    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Application.DisplayAlerts = True
    
    'Open main DTR
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=MainDTRReport
    Application.DisplayAlerts = True
    
    Range("I7").Select
    Range(Selection, Selection.End(xlDown)).Select
    CountDTRTemp = Selection.Cells.Count + 6
    Range("A7:R" & CountDTRTemp).Select
    Selection.Copy
    
    Application.DisplayAlerts = False
    ActiveWindow.Close SaveChanges:=True
    Application.DisplayAlerts = True

    'Select hidden Main
    DTRHTemplate1 = Sheets("main").Range("S8").Value

    Application.DisplayAlerts = False
    Workbooks.Open Filename:=DTRHTemplate1
    Application.DisplayAlerts = True
    
    Range("B7:R" & CountDTRTemp).Select
    ActiveSheet.Paste
    
    Range("A7:A" & CountDTRTemp).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
    
    Range("T7:V" & CountDTRTemp).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
  
    'General functions
    ValidForRowSix = Range("C6").Value
    If ValidForRowSix = "Classification" Then
        Range("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A7").Select
        ActiveCell.FormulaR1C1 = "=TRIM(RC[4])&TEXT(TRIM(RC[1]),""ddmmyyyy"")"
        
        'CountDTRTemp = Range("D7", "D10").Rows.Count
        Range("A7").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("A8:" & "A" & CountDTRTemp).Select
        ActiveSheet.Paste

    End If
    
    
    'Time IN
    Range("T7").Select
 
     'ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-16]=""Regular Working Day"",RC[-9]="""",RC[-2]=""""),""No Time Entry"",IF(AND(RC[-16]=""Rest Day"",RC[-9]=""""),""OFF"",(IF(AND(RC[-16]=""Regular Working Day"",RC[-9]<>""""),SUBSTITUTE(RC[-9],RIGHT(RC[-9],2),"":00 ""&RIGHT(RC[-9],2)),IF(AND(RC[-16]=""Rest Day"",RC[-9]<>""""),SUBSTITUTE(RC[-9],RIGHT(RC[-9],2),"":00 ""&RIGHT(RC[-9],2)),RC[-2])))))"
    
     ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-16]=""Regular Working Day"",RC[-9]="""",RC[-2]=""""),""No Time Entry"",IF(AND(RC[-16]=""Rest Day"",RC[-9]=""""),""OFF"",(IF(AND(RC[-16]=""Regular Working Day"",RC[-9]<>""""),SUBSTITUTE(RC[-9],RIGHT(RC[-9],2),"":00 ""&RIGHT(RC[-9],2)),IF(AND(RC[-16]=""Rest Day"",RC[-9]<>""""),SUBSTITUTE(RC[-9],RIGHT(RC[-9],2),"":00 ""&RIGHT(RC[-9],2)),IF(AND(RC[-9]<>""""," & _
        ".5),RC[-9],RC[-2]))))))"
    
    Range("T" & CountDTRTemp).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    
    'Time OUT
     Range("U7").Select
    'ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-17]=""Regular Working Day"",RC[-9]="""",RC[-3]=""""),""No Time Entry"",IF(AND(RC[-17]=""Rest Day"",RC[-9]=""""),""OFF"",(IF(AND(RC[-17]=""Regular Working Day"",RC[-9]<>""""),SUBSTITUTE(RC[-9],RIGHT(RC[-9],2),"":00 ""&RIGHT(RC[-9],2)),IF(AND(RC[-17]=""Rest Day"",RC[-9]<>""""),SUBSTITUTE(RC[-9],RIGHT(RC[-9],2),"":00 ""&RIGHT(RC[-9],2)),RC[-3])))))"
     ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-17]=""Regular Working Day"",RC[-9]="""",RC[-3]=""""),""No Time Entry"",IF(AND(RC[-17]=""Rest Day"",RC[-9]=""""),""OFF"",(IF(AND(RC[-17]=""Regular Working Day"",RC[-9]<>""""),SUBSTITUTE(RC[-9],RIGHT(RC[-9],2),"":00 ""&RIGHT(RC[-9],2)),IF(AND(RC[-17]=""Rest Day"",RC[-9]<>""""),SUBSTITUTE(RC[-9],RIGHT(RC[-9],2),"":00 ""&RIGHT(RC[-9],2)),IF(AND(RC[-9]<>""""," & _
        ".5),RC[-9],RC[-3]))))))"
    
    Range("U8").Select
    
    Range("U" & CountDTRTemp).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    
    'Close DTR
    ActiveWorkbook.Save
    ActiveWindow.Close
    
 
    Call Attendance(DateInputStart, DateInputEnd)
    

End Sub
