Attribute VB_Name = "Fintech"

Sub FintechEFTchangeDate()
Attribute FintechEFTchangeDate.VB_ProcData.VB_Invoke_Func = "F\n14"
    
    Columns("I:I").ColumnWidth = 17
    
    totrow = 1
    r = 2
    ic = 9      'Column I
    TC = 20     'Coumnn T
    
    Cells(r, ic).Select
    
    'Count how many rows
    Do While ActiveCell <> ""
    
        totrow = totrow + 1
        
        'Prepare to go to next row
        r = r + 1
        
        'Go to next cell
        Cells(r, ic).Select
    
    Loop
    
'    TotRowsQuestion = MsgBox("A total of " & totrow & " rows have been counted. Is this correct?", vbYesNo, "Total Rows")
'
'    'Ask User if there was a weekend
'    Dim AnswerYes As String
'    Dim AnswerNo As String
'
'    If TotRowsQuestion = vbYes Then
    
    WeekEnd = MsgBox("Was there a weekend between the Bank Activity Date and the EFT Pocessing Date?", vbYesNo, "Weekend")
    
    'Input formulat into T2
    Cells(2, TC).Select
    If WeekEnd = vbYes Then
        ActiveCell.FormulaR1C1 = "=RC[-11]+3"
    Else
        ActiveCell.FormulaR1C1 = "=RC[-11]+1"
    End If

    'Copy formula and paste to rest of Column T
    Cells(2, TC).Copy
    Range(Cells(3, TC), Cells(totrow, TC)).Select
    ActiveSheet.Paste
    
    'Copy Column T and PasteSpecial to Column I
    Range(Cells(2, TC), Cells(totrow, TC)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Cells(2, ic), Cells(totrow, ic)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'Delete Column T
    Range("T2:T54").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    
    'Saves File for DRSA
    Dim FileName As String
    Dim FilePath As String
    
'     File = ActiveWorkbook.Name
'    FileName = Split(File, ".")
'    FileNameParts = Split(FileName(0), " ")
'    StartDate = FileNameParts(10)
'    EndDate = FileNameParts(12)
    
    filename_ = Split(Cells(2, "i").Value, "/")
    If filename_(0) < 10 Then
        filename_MM = "0" & filename_(0)
    Else
        filename_MM = filename_(0)
    End If
    If filename_(1) < 10 Then
        filename_DD = "0" & filename_(1)
    Else
        filename_DD = filename_(1)
    End If
    filename_YY = Mid(filename_(2), 3, 2)
    FileName = "Fintech " & filename_MM & "." & filename_DD & "." & filename_YY
    FilePath = "\\Server\f\Accounting\S.Mitchell\Fintech\"
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:=FilePath & FileName & ".csv", FileFormat:=xlCSV
    Application.DisplayAlerts = True
    
End Sub
