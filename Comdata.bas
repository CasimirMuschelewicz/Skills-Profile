Attribute VB_Name = "Comdata"
Sub ComdataFileFormat()

    'Opens all necessary files
    Dim wb As Workbook
    Dim FolderPath As String
    Dim FilePath As String
    Dim MonthFile As String
    Dim StartDate As String
    Dim EndDate As String
    
    FolderPath = "C:\Users\smitchell\Desktop\Outlook Attachments\Comdata\"
    FilePath = Dir(FolderPath & "*Comdata Merchant Transaction Detail*" & ".xls")
    Do While FilePath <> ""
        Set wb = Workbooks.Open(FolderPath & FilePath)
        FilePath = Dir
    Loop
    ActiveSheet.Name = "Report"
    
    file = ActiveWorkbook.Name
    FileName = Split(file, ".")
    FileNameParts = Split(FileName(0), " ")
    StartDate = FileNameParts(10)
    EndDate = FileNameParts(12)
    
    If Len(StartDate) = 4 Then
        StartDate = "0" & Mid(StartDate, 1, 1) & ".0" & Mid(StartDate, 2, 1) & "." & Mid(StartDate, 3, 2)
    ElseIf Len(StartDate) = 5 Then
        StartDate = "0" & Mid(StartDate, 1, 1) & "." & Mid(StartDate, 2, 2) & "." & Mid(StartDate, 4, 2)
    Else
        StartDate = Mid(StartDate, 1, 2) & "." & Mid(StartDate, 3, 2) & "." & Mid(StartDate, 5, 2)
    End If
    
    If Len(EndDate) = 4 Then
        EndDate = "0" & Mid(EndDate, 1, 1) & ".0" & Mid(EndDate, 2, 1) & "." & Mid(EndDate, 3, 2)
    ElseIf Len(EndDate) = 5 Then
        EndDate = "0" & Mid(EndDate, 1, 1) & "." & Mid(EndDate, 2, 2) & "." & Mid(EndDate, 4, 2)
    Else
        EndDate = Mid(EndDate, 1, 2) & "." & Mid(EndDate, 3, 2) & "." & Mid(EndDate, 5, 2)
    End If
    
'    StartMM = Mid(StartDate, 1, 2)
'    StartDD = Mid(StartDate, 4, 2)
'    StartYY = Mid(StartDate, 7, 2)
'
'    EndMM = Mid(EndDate, 1, 2)
'    EndDD = Mid(EndDate, 4, 2)
'    EndYY = Mid(EndDate, 7, 2)
     
    'Formatting
    rstep = 8
    Range("A" & rstep).Select
    Do While ActiveCell <> ""
        rstep = rstep + 1
        Range("A" & rstep).Select
    Loop
    totrow_11 = rstep
    
    initialrow_17 = totrow_11 + 4
    rstep = totrow_11 + 4
    Range("A" & rstep).Select
    Do While ActiveCell <> ""
        rstep = rstep + 1
        Range("A" & rstep).Select
    Loop
    totrow_17 = rstep
    
    'Tony Format
    'Site 11
    Range("B8", "D" & totrow_11).Select
    Selection.Delete shift:=xlToLeft
    Range("D8", "T" & totrow_11).Select
    Selection.Delete shift:=xlToLeft
    Range("E8", "F" & totrow_11).Select
    Selection.Delete shift:=xlToLeft
    Range("F8", "G" & totrow_11).Select
    Selection.Delete shift:=xlToLeft
    Range("G8", "G" & totrow_11).Select
    Selection.Delete shift:=xlToLeft
    'Site 17
    Range("B" & initialrow_17, "D" & totrow_17).Select
    Selection.Delete shift:=xlToLeft
    Range("D" & initialrow_17, "T" & totrow_17).Select
    Selection.Delete shift:=xlToLeft
    Range("E" & initialrow_17, "F" & totrow_17).Select
    Selection.Delete shift:=xlToLeft
    Range("F" & initialrow_17, "G" & totrow_17).Select
    Selection.Delete shift:=xlToLeft
    Range("G" & initialrow_17, "G" & totrow_17).Select
    Selection.Delete shift:=xlToLeft
    Range("A1").Select
    
    'Tony File Save As
    FileName_Comdata_Tony = "Comdata " & StartDate & " to " & EndDate & "_Tony"
    FilePath_Comdata_Tony = "\\Server\f\Accounting\S.Mitchell\Comdata\Tony\" & FileName_Comdata_Tony & ".xlsx"
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:=FilePath_Comdata_Tony, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    'Devin Format
    'Site 11
    Range("E8", "F" & totrow_11).Select
    Selection.Delete shift:=xlToLeft
    'Site 17
    Range("E" & initialrow_17, "F" & totrow_17).Select
    Selection.Delete shift:=xlToLeft
    Range("A1").Select
    
    'Devin File Save As
    FileName_Comdata_Devin = "Comdata " & StartDate & " to " & EndDate
    FilePath_Comdata_Devin = "\\Server\f\Accounting\S.Mitchell\Comdata\Devin\" & FileName_Comdata_Devin & ".xlsx"
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:=FilePath_Comdata_Devin, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
End Sub

Sub Comdata_DRSA_Pivot()
    
    '-------------------------------------------------------------------------------------------------------
    'DRSA
    '--------------------------------------------------------
    Sheets(2).Select
    ActiveSheet.Name = "DRSA"
    
    Range("A1").Select
    cellcontains = Selection.Text
    
    If cellcontains = "-1" Then
        
        Columns("S:S").Delete (xlToLeft)
        totrow = ActiveSheet.UsedRange.Rows.Count
    
        Columns("A:A").Select
        Selection.Delete shift:=xlToLeft
        Range("A1:C1").Select
        Selection.Cut
        Range("A" & totrow + 1, "C" & totrow + 1).Select
        ActiveSheet.Paste
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "ID"
        Range("B1").Select
        ActiveCell.FormulaR1C1 = "Description"
        Range("C1").Select
        ActiveCell.FormulaR1C1 = "Comdata"
        Columns("A:A").ColumnWidth = 10
        Columns("B:B").ColumnWidth = 30
        Columns("C:C").ColumnWidth = 20
        Columns("C:C").Select
        Selection.Style = "Currency"
        Range("A1:C18").Select
        Application.CutCopyMode = False
        ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$C$18"), , xlYes).Name = _
            "Table1"
        Range("Table1[#All]").Select
        ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
        Range("A" & totrow + 1, "C" & totrow + 1).Select
        Selection.Font.Bold = True
        
    Else
        MsgBox ("Error: DRSA not in correct format")
        Exit Sub
    End If
        
    '-------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------
    'Pivot Tables
    '--------------------------------------------------------
        
    'Finds Food-N-Fun #11 and saves range
    Worksheets("Report").Select
    r = 1
    Cells(r, 1).Select
    Do Until ActiveCell.Value = "FOOD-N-FUN #11"
        r = r + 1
        Cells(r, 1).Select
    Loop
    r = r + 2
    FRow11 = r
    Do Until ActiveCell = ""
        r = r + 1
        Cells(r, 1).Select
    Loop
    LRow11 = r - 1
    Dim Table11 As Range
    Set Table11 = Sheets("Report").Range(Cells(FRow11, 1), Cells(LRow11, 4))
    
    
    'Finds Food-N-Fun #17 and saves range
    r = LRow11
    Cells(r, 1).Select
    Do Until ActiveCell.Value = "FOOD-N-FUN #17"
        r = r + 1
        Cells(r, 1).Select
    Loop
    r = r + 2
    FRow17 = r
    Do Until ActiveCell = ""
        r = r + 1
        Cells(r, 1).Select
    Loop
    LRow17 = r - 1
    Dim Table17 As Range
    Set Table17 = Sheets("Report").Range(Cells(FRow17, 1), Cells(LRow17, 4))
    
    'Creates new sheets
    Dim ws11 As Worksheet
    Table11.Copy
    Set ws11 = Sheets.Add
    ws11.Name = "Table 11"
    ActiveSheet.Paste Cells(1, 1)
    
    Dim ws17 As Worksheet
    Table17.Copy
    Set ws17 = Sheets.Add
    ws17.Name = "Table 17"
    ActiveSheet.Paste Cells(1, 1)
       
       

    '-------------------------------------------------------------------------------------------------------
    'Site 11
    '-------------------------------------------------------------------------------------------------------

    Dim PTable11Date As PivotTable    'Refers to Pivot Table
    Dim PTable11Time As PivotTable
    Dim PCache11 As PivotCache    'Holds the Data the Pivot Table refers to
    Dim PRange11 As Range         'The range of raw data
    Dim PSheet11 As Worksheet     'The sheet with the pivot table
    Dim DSheet11 As Worksheet     'The sheet with the data
    Dim LR11 As Long              'Last Row
    Dim LC11 As Long              'Last Column
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Worksheets.Add before:=ActiveSheet ' This will add new worksheet
    ActiveSheet.Name = "FOOD-N-FUN #11" ' This will rename the worksheet as "Pivot Sheet"
    On Error GoTo 0
    
    Set PSheet11 = Worksheets("FOOD-N-FUN #11")
    Set DSheet11 = Worksheets("Table 11")
    'Find Last used row and column in data sheet
    LR11 = DSheet11.Cells(Rows.Count, 1).End(xlUp).row
    LC11 = DSheet11.Cells(1, Columns.Count).End(xlToLeft).Column
    'Set the pivot table data range
    Set PRange11 = DSheet11.Cells(1, 1).Resize(LR11, LC11)
    'Set pivot cahe
    Set PCache11 = ActiveWorkbook.PivotCaches.Create(xlDatabase, SourceData:=PRange11)
    
    
'Create Pivot Table Totals by Date-----------------------------------------
    'Create blank pivot table
    Set PTable11Date = PCache11.CreatePivotTable(TableDestination:=PSheet11.Cells(3, 2), _
                       TableName:="Totals by Date")
    'Insert country to Row Filed
    With PSheet11.PivotTables("Totals by Date").PivotFields("Invoice Date")
    .Orientation = xlRowField
    .Position = 1
    End With
    'Insert Segment to Column Filed & position 1
    With PSheet11.PivotTables("Totals by Date").PivotFields("Invoice Total")
    .Orientation = xlDataField
    .Position = 1
    End With
    'Format Pivot Table
    PSheet11.PivotTables("Totals by Date").ShowTableStyleRowStripes = False
    PSheet11.PivotTables("Totals by Date").TableStyle2 = "PivotStyleMedium9"
    Cells(2, 2).Value = "Totals by Date"
    With Range(Cells(2, 2), Cells(2, 3))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
        .Font.Bold = True
        .Font.Size = 16
    End With
    Columns("c:c").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
'Create Pivot Table Totals by Date & Time-------------------------------------
    'Create blank pivot table
    Set PTable11Time = PCache11.CreatePivotTable(TableDestination:=PSheet11.Cells(3, 6), _
                       TableName:="Totals by Date & Time")
    'Insert country to Row Filed
    With PSheet11.PivotTables("Totals by Date & Time").PivotFields("Invoice Date")
    .Orientation = xlRowField
    .Position = 1
    End With
    'Insert country to Row Filed
    With PSheet11.PivotTables("Totals by Date & Time").PivotFields("Inv Time")
    .Orientation = xlRowField
    .Position = 2
    End With
    'Insert Segment to Column Filed & position 1
    With PSheet11.PivotTables("Totals by Date & Time").PivotFields("Invoice Total")
    .Orientation = xlDataField
    .Position = 1
    End With
    'Format Pivot Table
    PSheet11.PivotTables("Totals by Date & Time").ShowTableStyleRowStripes = False
    PSheet11.PivotTables("Totals by Date & Time").TableStyle2 = "PivotStyleMedium9"
    Cells(2, 6).Value = "Totals by Date & Time"
    With Range(Cells(2, 6), Cells(2, 7))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
        .Font.Bold = True
        .Font.Size = 16
    End With
    Columns("g:g").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ActiveWindow.DisplayGridlines = False
    
    Columns("A:A").ColumnWidth = 70
    Columns("D:D").ColumnWidth = 10
    Columns("E:E").ColumnWidth = 10
    Columns("H:H").ColumnWidth = 15
    
'    'Check Boxes
'    r = 4
'    r2 = 4
'    c = 0
'    Cells(r, 6).Select
'    Do While ActiveCell <> ""
'        If TypeName(Cells(r, 6).Value) = "Date" Then
'            ActiveSheet.Checkboxes.Add(Cells(r, "h").Left + 10, _
'                                       Cells(r, "h").Top, _
'                                       Cells(r, "h").Width - 15, _
'                                       Cells(r, "h").Height).Select
'            With Selection
'                .Caption = "Reconciled"
'                .LinkedCell = "i" & r
'                .Display3DShading = False
'            End With
'
'            Cells(r2, "d").Select
'            With Selection
'                .Value = "=IF(I" & r & "=TRUE,""REC"","""")"
'                .HorizontalAlignment = xlCenter
'                .VerticalAlignment = xlCenter
'                .Font.color = vbRed
'            End With
'            r2 = r2 + 1
'            c = c + 1
'            If c = 7 Then
'                Exit Do
'            End If
'        End If
'        r = r + 1
'        Cells(r, 6).Select
'    Loop
'    Columns("i:i").Hidden = True
    
    Columns("A:A").Delete
    
    Set i = Nothing
    Set r = Nothing
    Set r2 = Nothing
    Set c = Nothing
    
    Worksheets("FOOD-N-FUN #11").Select

    
    
    
    '-------------------------------------------------------------------------------------------------------
    'Site 17
    '-------------------------------------------------------------------------------------------------------

    Dim PTable17Date As PivotTable    'Refers to Pivot Table
    Dim PTable17Time As PivotTable
    Dim PCache17 As PivotCache    'Holds the Data the Pivot Table refers to
    Dim PRange17 As Range         'The range of raw data
    Dim PSheet17 As Worksheet     'The sheet with the pivot table
    Dim DSheet17 As Worksheet     'The sheet with the data
    Dim LR17 As Long              'Last Row
    Dim LC17 As Long              'Last Column
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Worksheets.Add after:=ActiveSheet ' This will add new worksheet
    ActiveSheet.Name = "FOOD-N-FUN #17" ' This will rename the worksheet as "Pivot Sheet"
    On Error GoTo 0
    
    Set PSheet17 = Worksheets("FOOD-N-FUN #17")
    Set DSheet17 = Worksheets("Table 17")
    'Find Last used row and column in data sheet
    LR17 = DSheet17.Cells(Rows.Count, 1).End(xlUp).row
    LC17 = DSheet17.Cells(1, Columns.Count).End(xlToLeft).Column
    'Set the pivot table data range
    Set PRange17 = DSheet17.Cells(1, 1).Resize(LR17, LC17)
    'Set pivot cahe
    Set PCache17 = ActiveWorkbook.PivotCaches.Create(xlDatabase, SourceData:=PRange17)
    
'Create Pivot Table Totals by Date-----------------------------------------
    'Create blank pivot table
    Set PTable17Date = PCache17.CreatePivotTable(TableDestination:=PSheet17.Cells(3, 2), _
                       TableName:="Totals by Date")
    'Insert country to Row Filed
    With PSheet17.PivotTables("Totals by Date").PivotFields("Invoice Date")
    .Orientation = xlRowField
    .Position = 1
    End With
    'Insert Segment to Column Filed & position 1
    With PSheet17.PivotTables("Totals by Date").PivotFields("Invoice Total")
    .Orientation = xlDataField
    .Position = 1
    End With
    'Format Pivot Table
    PSheet17.PivotTables("Totals by Date").ShowTableStyleRowStripes = False
    PSheet17.PivotTables("Totals by Date").TableStyle2 = "PivotStyleMedium9"
    Cells(2, 2).Value = "Totals by Date"
    With Range(Cells(2, 2), Cells(2, 3))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
        .Font.Bold = True
        .Font.Size = 16
    End With
    Columns("c:c").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
'Create Pivot Table Totals by Date & Time-------------------------------------
    'Create blank pivot table
    Set PTable17Time = PCache17.CreatePivotTable(TableDestination:=PSheet17.Cells(3, 6), _
                       TableName:="Totals by Date & Time")
    'Insert country to Row Filed
    With PSheet17.PivotTables("Totals by Date & Time").PivotFields("Invoice Date")
    .Orientation = xlRowField
    .Position = 1
    End With
    'Insert country to Row Filed
    With PSheet17.PivotTables("Totals by Date & Time").PivotFields("Inv Time")
    .Orientation = xlRowField
    .Position = 2
    End With
    'Insert Segment to Column Filed & position 1
    With PSheet17.PivotTables("Totals by Date & Time").PivotFields("Invoice Total")
    .Orientation = xlDataField
    .Position = 1
    End With
    'Format Pivot Table
    PSheet17.PivotTables("Totals by Date & Time").ShowTableStyleRowStripes = False
    PSheet17.PivotTables("Totals by Date & Time").TableStyle2 = "PivotStyleMedium9"
    Cells(2, 6).Value = "Totals by Date & Time"
    With Range(Cells(2, 6), Cells(2, 7))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
        .Font.Bold = True
        .Font.Size = 16
    End With
    Columns("g:g").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ActiveWindow.DisplayGridlines = False
    
    Columns("A:A").ColumnWidth = 70
    Columns("D:D").ColumnWidth = 10
    Columns("E:E").ColumnWidth = 10
    Columns("H:H").ColumnWidth = 15
    
'    'Check Boxes
'    r = 4
'    r2 = 4
'    c = 0
'    Cells(r, 6).Select
'    Do While ActiveCell <> ""
'        If TypeName(Cells(r, 6).Value) = "Date" Then
'            ActiveSheet.Checkboxes.Add(Cells(r, "h").Left + 10, _
'                                       Cells(r, "h").Top, _
'                                       Cells(r, "h").Width - 15, _
'                                       Cells(r, "h").Height).Select
'            With Selection
'                .Caption = "Reconciled"
'                .LinkedCell = "i" & r
'                .Display3DShading = False
'            End With
'
'            Cells(r2, "d").Select
'            With Selection
'                .Value = "=IF(I" & r & "=TRUE,""REC"","""")"
'                .HorizontalAlignment = xlCenter
'                .VerticalAlignment = xlCenter
'                .Font.color = vbRed
'            End With
'            r2 = r2 + 1
'            c = c + 1
'            If c = 7 Then
'                Exit Do
'            End If
'        End If
'        r = r + 1
'        Cells(r, 6).Select
'    Loop
'    Columns("i:i").Hidden = True
    
    Columns("A:A").Delete
    
    oldvalue = Application.DisplayAlerts
    Application.DisplayAlerts = False
    DSheet11.Delete
    DSheet17.Delete
    Application.DisplayAlerts = oldvalue
    PSheet11.Activate
    Range("a1").Select
    
    ActiveWorkbook.Save
    
End Sub

Sub ComdataDRSA()

    Range("A1").Select
    cellcontains = Selection.Text
    
    If cellcontains = "-1" Then
        
        Columns("S:S").Delete (xlToLeft)
        
        totrow = 0
        r = 1
        
        'Count how many rows
        Do While ActiveCell <> ""
            totrow = totrow + 1
            'Prepare to go to next row
            r = r + 1
            'Go to next cell
            Range("A" & r).Select
        Loop
    
        Columns("A:A").Select
        Selection.Delete shift:=xlToLeft
        Range("A1:C1").Select
        Selection.Cut
        Range("A" & totrow + 1, "C" & totrow + 1).Select
        ActiveSheet.Paste
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "ID"
        Range("B1").Select
        ActiveCell.FormulaR1C1 = "Description"
        Range("C1").Select
        ActiveCell.FormulaR1C1 = "Comdata"
        Columns("A:A").ColumnWidth = 10
        Columns("B:B").ColumnWidth = 30
        Columns("C:C").ColumnWidth = 20
        Columns("C:C").Select
        Selection.Style = "Currency"
        Range("A1:C18").Select
        Application.CutCopyMode = False
        ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$C$18"), , xlYes).Name = _
            "Table1"
        Range("Table1[#All]").Select
        ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
        Range("A" & totrow + 1, "C" & totrow + 1).Select
        Selection.Font.Bold = True
        
        'User inputs dates
        StartDate = InputBox("Enter Start Date: ", "Start Date")
        EndDate = InputBox("Enter End Date: ", "End Date")
        
        Dim FileName_Comdata_DRSA As String
        Dim FilePath_Comdata_DRSA As String
        
        'Saves File for DRSA
        FileName_Comdata_DRSA = "Comdata " & StartDate & " to " & EndDate & "_DRSA"
        FilePath_Comdata_DRSA = "\\Server\f\Accounting\S.Mitchell\Comdata\Devin\" & _
                                FileName_Comdata_DRSA & ".xlsx"
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs FileName:=FilePath_Comdata_DRSA, FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        
    Else
        MsgBox ("Error: DRSA not in correct format")
    End If
    
End Sub

Sub ComdataFileFormat_Old()

    Dim MonthFile As String
    Dim StartDate As String
    Dim EndDate As String
    
    'User inputs dates
    Do
        YearFile = InputBox("Enter the Year (YYYY).", "Year")
    Loop Until Len(YearFile) <> 0 Or StrPtr(YearFile) = 0
    If StrPtr(YearFile) = 0 Then
        MsgBox ("User canceled")
    Else
        Do
            MonthFile = InputBox("Enter the Month (MM).", "Month")
        Loop Until Len(MonthFile) <> 0 Or StrPtr(MonthFile) = 0
        If StrPtr(MonthFile) = 0 Then
            MsgBox ("User canceled")
        Else
            Do
                StartDate = InputBox("Enter Start Date MMDDYY: ", "Start Date")
            Loop Until Len(StartDate) <> 0 Or StrPtr(StartDate) = 0
            If StrPtr(StartDate) = 0 Then
                MsgBox ("User canceled")
            ElseIf StartDate = vbNullString Then
                        
            Else
                Do
                    EndDate = InputBox("Enter End Date MMDDYY: ", "End Date")
                Loop Until Len(EndDate) <> 0 Or StrPtr(EndDate) = 0
                If StrPtr(EndDate) = 0 Then
                    MsgBox ("User canceled")
                ElseIf EndDate = vbNullString Then
                            
                Else
                    
                    rstep = 8
                    Range("A" & rstep).Select
                    Do While ActiveCell <> ""
                        rstep = rstep + 1
                        Range("A" & rstep).Select
                    Loop
                    totrow_11 = rstep
                    
                    initialrow_17 = totrow_11 + 4
                    rstep = totrow_11 + 4
                    Range("A" & rstep).Select
                    Do While ActiveCell <> ""
                        rstep = rstep + 1
                        Range("A" & rstep).Select
                    Loop
                    totrow_17 = rstep
                    
                    'Tony Format
                    'Site 11
                    Range("B8", "D" & totrow_11).Select
                    Selection.Delete shift:=xlToLeft
                    Range("D8", "T" & totrow_11).Select
                    Selection.Delete shift:=xlToLeft
                    Range("E8", "F" & totrow_11).Select
                    Selection.Delete shift:=xlToLeft
                    Range("F8", "G" & totrow_11).Select
                    Selection.Delete shift:=xlToLeft
                    Range("G8", "G" & totrow_11).Select
                    Selection.Delete shift:=xlToLeft
                    'Site 17
                    Range("B" & initialrow_17, "D" & totrow_17).Select
                    Selection.Delete shift:=xlToLeft
                    Range("D" & initialrow_17, "T" & totrow_17).Select
                    Selection.Delete shift:=xlToLeft
                    Range("E" & initialrow_17, "F" & totrow_17).Select
                    Selection.Delete shift:=xlToLeft
                    Range("F" & initialrow_17, "G" & totrow_17).Select
                    Selection.Delete shift:=xlToLeft
                    Range("G" & initialrow_17, "G" & totrow_17).Select
                    Selection.Delete shift:=xlToLeft
                    Range("A1").Select
                    
                    'Tony File Save As
                    FileName_Comdata_Tony = "Comdata " & StartDate & " to " & EndDate & "_Tony"
                    FilePath_Comdata_Tony = "\\Server\f\Accounting\S.Mitchell\Comdata\Tony\" & YearFile & "\" & MonthFile & "\" & FileName_Comdata_Tony & ".xlsx"
                    Application.DisplayAlerts = False
                    ActiveWorkbook.SaveAs FileName:=FilePath_Comdata_Tony, FileFormat:=xlOpenXMLWorkbook
                    Application.DisplayAlerts = True
                    
                    'Devin Format
                    Range("E8", "F" & totrow_11).Select
                    Selection.Delete shift:=xlToLeft
                    Range("E" & initialrow_17, "F" & totrow_17).Select
                    Selection.Delete shift:=xlToLeft
                    Range("A1").Select
                    
                    'Devin File Save As
                    FileName_Comdata_Devin = "Comdata " & StartDate & " to " & EndDate
                    FilePath_Comdata_Devin = "\\Server\f\Accounting\S.Mitchell\Comdata\Devin\" & FileName_Comdata_Devin & ".xlsx"
                    Application.DisplayAlerts = False
                    ActiveWorkbook.SaveAs FileName:=FilePath_Comdata_Devin, FileFormat:=xlOpenXMLWorkbook
                    Application.DisplayAlerts = True
                End If
            End If
        End If
    End If
    
End Sub
