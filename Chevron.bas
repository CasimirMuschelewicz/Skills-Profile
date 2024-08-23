Attribute VB_Name = "Chevron"
Sub Chevron_Pre_Merge()

    Worksheets("Sheet1").Select
    
    Range("A1").Select
    cellcontains = Selection.Text
    If cellcontains = "-1" Then
        r = 1
        'Count how many rows
        Do While ActiveCell <> ""
            'Prepare to go to next row
            r = r + 1
            'Go to next cell
            Range("A" & r).Select
        Loop
        totrow = r - 1
       
        Columns("A:A").Select
        Selection.Delete shift:=xlToLeft
        Columns("A:E").Select
        Selection.ColumnWidth = 15
        Range("A1:E1").Select
        Selection.Cut
        Range("A" & totrow + 2, "D" & totrow + 2).Select 'Puts Grand Total line 2 rows below data
        ActiveSheet.Paste
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "ID"
        Range("B1").Select
        ActiveCell.FormulaR1C1 = "Description"
        Range("C1").Select
        ActiveCell.FormulaR1C1 = "P97"
        Range("D1").Select
        ActiveCell.FormulaR1C1 = "Giftcards"
        Range("E1").Select
        ActiveCell.FormulaR1C1 = "Total EMS"
        Range("A1:E1").Select
        Selection.Font.Bold = True
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        Columns("C:C").Select
        Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        'Inserts corresponding site numbers for every row
        Range("A2", "A" & totrow).Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.FormulaR1C1 = "=R[-1]C"
        Range("A2", "A" & totrow).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("A2", "A" & totrow).Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        
        'Deletes all sub total rows
        step = 2
        Cells(step, 2).Select
        Do While ActiveCell <> ""
            If InStr(1, ActiveCell.Value, "Food") <> 0 Then
                ActiveCell.EntireRow.Delete shift:=xlUp
            Else
                '...
            End If
            step = step + 1
            Cells(step, 2).Select
        Loop
        
        Dim i As Integer
        Dim Day As String
        
        row = 2
        col = 2
        Cells(row, col).Select
        
        'Changes format of dates to match DRRF date format
        Do While ActiveCell <> ""
            Day = Cells(row, col).Value
            Day = Trim(Day)                         'Removes all spaces from cell
            If InStr(Day, "0") = 1 Then
                If InStr(4, Day, "0") = 4 Then
                    Day = Replace(Day, "0", "", 1, 2)
                Else
                    Day = Replace(Day, "0", "", 1, 1)
                End If
            Else
                If InStr(4, Day, "0") = 4 Then
                    i = 4
                    Day = Mid(Day, 1, i - 1) & Replace(Day, "0", "", i, 1)
                Else
                    'No zeros to replace
                End If
            End If
            Cells(row, col).Value = Day
            row = row + 1
            Cells(row, col).Select                   'Selects next cell to be cleaned
        Loop
        
        Columns("B:B").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        'Count rows after deleting subtotal rows
        r = 2
        Range("A" & r).Select
        Do While ActiveCell <> ""
            r = r + 1
            Range("A" & r).Select
        Loop
        totrow = r - 1
        EMSinitialRow = totrow + 1
        EMStotRow = EMSinitialRow + totrow - 2
        
        'duplicates table below original
        Range(Cells(2, 1), Cells(totrow, 6)).Select
        Selection.Copy
        Range(Cells(EMSinitialRow, 1), Cells(EMSinitialRow, 6)).Select
        Selection.Insert shift:=xlDown
        
        'Deletes duplicate P97, Giftcard, and original EMS columns
        Range(Cells(EMSinitialRow, 4), Cells(EMStotRow, 5)).Select
        Selection.Delete shift:=xlToLeft
        Columns("F:F").ClearContents
        
        'Adds "P97" and "EMS" label to respective rows
        Range(Cells(2, 3), Cells(totrow, 3)).Value = "P97"
        Range(Cells(EMSinitialRow, 3), Cells(EMStotRow, 3)).Value = "EMS"
        
        'Copies P97 and EMS Table for DR
        Range(Cells(2, 1), Cells(EMStotRow, 4)).Select
        Selection.Copy
    
    Else
        MsgBox ("DRSA in wrong format.")
        Exit Sub
    End If

    '-------

    'Goes to DR sheet and counts total rows and pastes DRSA below it
    Worksheets("DR Transaction Forms Daily Summ").Select
    countDRRF = 1
    Range("A" & countDRRF).Select
    Do While ActiveCell <> ""
        countDRRF = countDRRF + 1
        Range("A" & countDRRF).Select
    Loop
    totrowDRRF = countDRRF - 1
    Range("A" & countDRRF + 1).PasteSpecial (xlPasteValues)
    
    'Formats DRRF table
    Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("A:A,C:C").ColumnWidth = 8
    Range("B:B,D:F").ColumnWidth = 15
    Cells.Select
    Selection.RowHeight = 15
    Range("A1", "C" & totrowDRRF).Select
    Selection.UnMerge
    Range("C1", "C" & totrowDRRF).Select
    Selection.Delete shift:=xlToLeft
    Selection.NumberFormat = "General"
    Range("A1:D1").Merge Across:=True
    Range("A2:D2").Merge Across:=True
    Range("A3:D3").Merge Across:=True
    Range("A4:D4").Merge Across:=True
        
    'Inserts 14 bank rows below each store table and deletes store title row
    row_step = 6
    Range("A" & row_step).Select
    Do While ActiveCell <> ""
        If InStr(1, ActiveCell.Value, "Food") <> 0 Then
            If Cells(ActiveCell.row, 4) = "" Then
                Selection.EntireRow.Delete shift:=xlUp
            Else
                Range("A" & row_step, "D" & row_step + 27).Select
                Selection.Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                row_step = row_step + 29
                Range("A" & row_step).Select
            End If
        Else
            row_step = row_step + 1
            Range("A" & row_step).Select
        End If
    Loop
    
    'Sorts DRSA by store, pairing P97 and EMS with same dates
    DRSAstartrow = row_step + 1
    DRSAtotrow = DRSAstartrow
    Range("A" & DRSAstartrow).Select
    Do While ActiveCell <> ""
        DRSAtotrow = DRSAtotrow + 1
        Range("A" & DRSAtotrow).Select
    Loop
    DRSAtotrow = DRSAtotrow - 1
    ActiveWorkbook.Worksheets("DR Transaction Forms Daily Summ").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("DR Transaction Forms Daily Summ").Sort.SortFields. _
        Add2 Key:=Range("A" & DRSAstartrow, "A" & DRSAtotrow), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DR Transaction Forms Daily Summ").Sort
        .SetRange Range("A" & DRSAstartrow, "D" & DRSAtotrow)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Copies specific DRSA section
    row_step = DRSAstartrow
    Range("A" & row_step).Select
    site = Selection.Value
    DRRFrow = 6
    Do While ActiveCell <> ""
        site = Selection.Value
        initialrow = row_step
        Do Until ActiveCell.Value <> site
            row_step = row_step + 1
            Range("A" & row_step).Select
        Loop
        Range("A" & initialrow, "D" & row_step - 1).Copy
        
        'Pastes Copied cells in desired location
        Range("A" & DRRFrow).Select
        Do While ActiveCell <> ""
            DRRFrow = DRRFrow + 1
            Range("A" & DRRFrow).Select
        Loop
        DRRFinitial = DRRFrow
        Do Until InStr(1, ActiveCell.Value, "Food") <> 0
            DRRFrow = DRRFrow + 1
            Range("A" & DRRFrow).Select
        Loop
        Range("A" & DRRFinitial, "D" & DRRFrow - 1).PasteSpecial (xlPasteValues)
        DRRFrow = DRRFrow + 1
        Range("A" & row_step).Select
    Loop
    Range("A" & DRSAstartrow, "D" & DRSAtotrow).Delete shift:=xlUp
    
    'Sorts each store's entries by date
    row_step = 6
    Range("A" & row_step).Select
    Do While InStr(1, ActiveCell.Value, "Report") = 0
        initial_row = row_step
        Range("B" & row_step).Value = CStr(Range("B" & row_step).Value)
        Do While InStr(1, ActiveCell.Value, "Food") = 0
            row_step = row_step + 1
            Range("B" & row_step).Value = CStr(Range("B" & row_step).Value)
            Range("A" & row_step).Select
        Loop
        Range("A" & initial_row, "D" & row_step - 1).Select
        Selection.Sort Key1:=Range("B" & initial_row), Order1:=xlAscending, _
                       Header:=xlNo
        row_step = row_step + 1
        Range("A" & row_step).Select
    Loop
    Columns("B:B").NumberFormat = "m/d/yyyy"
    
    'Deletes P97 and EMS blank cells
    rowstep = 6
    Range("D" & rowstep).Select
    Do While InStr(1, Range("A" & rowstep), "Report") = 0
        If ActiveCell = "" Then
            ActiveCell.EntireRow.Delete (xlUp)
        Else
            rowstep = rowstep + 1
        End If
        Range("D" & rowstep).Select
    Loop
    
    'Updates store totals with formulas & clears site-total batch numbers
    row_step = 6
    Range("A" & row_step).Select
    Do While InStr(1, Range("A" & row_step).Value, "Report") = 0
        initial_row = row_step
        Do While InStr(1, ActiveCell.Value, "Food") = 0
            row_step = row_step + 1
            Range("A" & row_step).Select
        Loop
        Range("c" & row_step).ClearContents
        Range("D" & row_step).Formula = "=SUM(D" & initial_row & ":D" & row_step - 1 & ")"
        row_step = row_step + 1
        Range("A" & row_step).Select
    Loop
    
    'Updates Grand Total with formula
    Dim firstAddress As String, c As Range, rALL As Range
    With Worksheets("DR Transaction Forms Daily Summ").Range("a1", "A" & row_step)
        Set c = .Find("Food", LookIn:=xlValues, LookAt:=xlPart)
        If Not c Is Nothing Then
            Set rALL = c
            firstAddress = c.Address
            Do
                Set rALL = Union(rALL, c)
                Worksheets("DR Transaction Forms Daily Summ").Range(c.Address).Activate
                Set c = .FindNext(c)

            Loop While Not c Is Nothing And c.Address <> firstAddress
        End If
        If Not rALL Is Nothing Then
            Cells(row_step, "d").Formula = "=SUM(" & rALL.Offset(, 3).Address & ")"
        End If
    End With
    
    'Adds a negative(-) sign to each entry
    rowstep = 6
    Range("d" & rowstep).Select
    Do While InStr(1, Range("a" & rowstep), "Report") = 0
        If InStr(1, Range("a" & rowstep), "Food") = 0 Then
            If ActiveCell.Value < 0 Then
                ActiveCell.Value = Abs(ActiveCell.Value)
            Else
                ActiveCell.Value = "-" & ActiveCell.Value
            End If
        End If
        rowstep = rowstep + 1
        Range("d" & rowstep).Select
    Loop
    
    'Formats D column to Accounting
    Range("D6", "D" & rowstep).Select
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    'Replaces "Food-N-Fun..." label with site number only
    rowstep = 6
    Range("A" & rowstep).Select
    Do While InStr(1, Range("A" & rowstep), "Report") = 0
        If InStr(1, ActiveCell.Value, "Food") <> 0 Then
            site = ActiveCell.Value
            site = Replace(site, "Food-N-Fun ", "", 1, 11)
            ActiveCell.Value = Mid(site, 1, 2) & Replace(site, "Total", "", 3)
            Selection.HorizontalAlignment = xlLeft
        End If
        rowstep = rowstep + 1
        Range("A" & rowstep).Select
    Loop
    
    'Sorts final DRRF by site and batch number
    Range("A6", "D" & rowstep).Select
    Selection.Sort Key1:=Range("a6", "a" & rowstep), Order1:=xlAscending, _
                   Key2:=Range("c6", "c" & rowstep), Order1:=xlAscending, _
                       Header:=xlNo
                       
    'Formats for merge with DTN
    Cells.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    ActiveWindow.DisplayGridlines = True
    Range("A1:D4").Select
    Selection.Delete shift:=xlUp
    Columns("B:B").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Columns("B:B").Select
    Selection.Insert shift:=xlToRight
    Selection.ColumnWidth = 15
    Columns("D:D").Select
    Selection.Copy
    Columns("B:B").Select
    ActiveSheet.Paste
    Columns("D:D").Select
    Selection.ClearContents
    
    'Finalizes format for Giftcards
    Worksheets("Sheet1").Select
    Worksheets("Sheet1").Name = "Giftcards"
    Columns("C:D").Delete shift:=xlToLeft
    totrow = Cells(Rows.Count, 1).End(xlUp).row
    Range("C2", "C" & totrow).SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete (xlUp)
    Range("a1").Select
    
    'Formats Append
    Dim ws_append As Worksheet
    For Each ws_append In ActiveWorkbook.Worksheets
        If ws_append.Name Like "*APPEND*" Then
            ws_append.Activate
            Exit For
        End If
    Next ws_append
    ActiveSheet.Name = "Append"
    ActiveSheet.ListObjects(1).TableStyle = ""
    ActiveSheet.ListObjects(1).Unlist
    Columns("A:B").Delete xlToLeft
    t = ActiveSheet.UsedRange.Rows.Count
    Range(Cells(1, 1), Cells(t, 5)).Select
    On Error Resume Next
    Selection.SpecialCells(xlCellTypeBlanks).Delete (xlUp)
    On Error GoTo 0

    'Renames DR and creates copy to format merged
    Worksheets("DR Transaction Forms Daily Summ").Select
    Worksheets("DR Transaction Forms Daily Summ").Name = "DRRF"
    Sheets("DRRF").Copy before:=Sheets("DRRF")
    ActiveSheet.Name = "Merged DRRF & DTN"
    
    'Formats DTN
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name Like "DTN*" Then
            ws.Activate
            Exit For
        End If
    Next ws
    
    'Formats Table
    ActiveSheet.ListObjects(1).TableStyle = ""
    ActiveSheet.ListObjects(1).Unlist
    Range("A1").EntireColumn.Delete (xlToLeft)
    
    'Replaces all site codes with store numbers that is ours
    t = ActiveSheet.UsedRange.Rows.Count
    Range("a2", "a" & t).Select
    'Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete (xlUp)
    rstep = 2
    Cells(rstep, "a").Select
    Do While rstep <> t + 1
        If ActiveCell = "109294" Or ActiveCell.Value = 2 Then
            ActiveCell.Value = 2
        ElseIf ActiveCell = "175002" Or ActiveCell.Value = 4 Then
            ActiveCell.Value = 4
        ElseIf ActiveCell = "203060" Or ActiveCell.Value = 7 Then
            ActiveCell.Value = 7
        ElseIf ActiveCell = "200239" Or ActiveCell.Value = 8 Then
            ActiveCell.Value = 8
        ElseIf ActiveCell = "201209" Or ActiveCell.Value = 10 Then
            ActiveCell.Value = 10
        ElseIf ActiveCell = "203062" Or ActiveCell.Value = 11 Then
            ActiveCell.Value = 11
        ElseIf ActiveCell = "205066" Or ActiveCell.Value = 13 Then
            ActiveCell.Value = 13
        ElseIf ActiveCell = "109348" Or ActiveCell.Value = 15 Then
            ActiveCell.Value = 15
        ElseIf ActiveCell = "205271" Or ActiveCell.Value = 16 Then
            ActiveCell.Value = 16
        ElseIf ActiveCell = "209207" Or ActiveCell.Value = 17 Then
            ActiveCell.Value = 17
        ElseIf ActiveCell = "208180" Or ActiveCell.Value = 18 Then
            ActiveCell.Value = 18
        ElseIf ActiveCell = "212523" Or ActiveCell.Value = 19 Then
            ActiveCell.Value = 19
        ElseIf ActiveCell = "372295" Or ActiveCell.Value = 20 Then
            ActiveCell.Value = 20
        ElseIf ActiveCell = "309462" Or ActiveCell.Value = 22 Then
            ActiveCell.Value = 22
        ElseIf ActiveCell = "108981" Or ActiveCell.Value = 24 Then
            ActiveCell.Value = 24
        ElseIf ActiveCell = "382364" Or ActiveCell.Value = 27 Then
            ActiveCell.Value = 27
        ElseIf ActiveCell = "357325" Or ActiveCell.Value = 29 Then
            ActiveCell.Value = 29
        ElseIf ActiveCell = "210685" Or ActiveCell.Value = 31 Then
            ActiveCell.Value = 31
        Else
            ActiveCell.ClearContents
        End If
        rstep = rstep + 1
        Cells(rstep, "a").Select
    Loop
    t = ActiveSheet.UsedRange.Rows.Count
    Range("a2", "a" & t).Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete (xlUp)
    rstep = 2
    Cells(rstep, "b").Select
    Do While ActiveCell <> ""
        txt = ActiveCell.Value
        ActiveCell.Value = Mid(txt, 1, 4)
        rstep = rstep + 1
        Cells(rstep, "b").Select
    Loop
    Range("a2", Cells(t, 1)).EntireRow.Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add2 Key:= _
        Range("A2:A" & t), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A1:S" & t)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Cells(1, 1).Select
    ActiveSheet.Copy before:=Sheets("DRRF")
    ActiveSheet.Name = "DTN"
    
    'Moves all entries on Append to their appropriate places
    Worksheets("Append").Select
    t = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To t
        Cells(i, 4).Select
        If Selection.Value = "DR/FNF" Or Selection.Value = "DR/TR" Then
            Selection.EntireRow.Copy
            Sheets("Merged DRRF & DTN").Select
            r = ActiveSheet.UsedRange.Rows.Count
            Cells(r + 1, 1).Select
            Selection.PasteSpecial xlPasteValues
        ElseIf Selection.Value = "DTN" Or Selection.Value = "DTN/TR" Then
            Selection.EntireRow.Copy
            Sheets("DTN").Select
            r = ActiveSheet.UsedRange.Rows.Count
            Cells(r + 1, 1).Select
            Selection.PasteSpecial xlPasteValues
        Else
            MsgBox ("Append Source not correct")
            Exit Sub
        End If
        Sheets("Append").Select
    Next i
    
    ' Sorts both Merged DRRF & DTN and DTN
    Sheets("Merged DRRF & DTN").Select
    t = ActiveSheet.UsedRange.Rows.Count
    Range("A2:A" & t).EntireRow.Select
    Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlNo, _
                   Key2:=Range("B2"), Order1:=xlAscending, Header:=xlNo
    Sheets("DTN").Select
    t = ActiveSheet.UsedRange.Rows.Count
    Range("A2:A" & t).EntireRow.Select
    Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlNo, _
                   Key2:=Range("B2"), Order1:=xlAscending, Header:=xlNo

'Saves File
    Dim StartDate As String
    Dim EndDate As String
    Dim FileName_Chev_DRSA As String
    Dim FilePath_Chev_DRSA As String
    
    StartDate = InputBox("Enter Start Date (mm.dd.yy): ", "Start Date")
    EndDate = InputBox("Enter End Date (mm.dd.yy): ", "End Date")
    
    'Saves File for DRSA
    FileName_Chev_DRSA = "Pre-Merge DRRF & DTN " & StartDate & " to " & EndDate
    FilePath_Chev_DRSA = "\\Server\f\Accounting\S.Mitchell\Chevron\DRRF_DRSA_DTN\Pre-Merge DRRF & DTN\" & FileName_Chev_DRSA & ".xlsx"
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:=FilePath_Chev_DRSA, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

End Sub

Sub Chevron_Merge()
    
    Sheets("DTN").Select
    r = 2
    Cells(r, 1).Select
    Do While ActiveCell <> ""
        
        DTNsite = Cells(r, 1).Value
        DTNbatch = Cells(r, 2).Value
        DTNamt = (Cells(r, 5).Value) * -1
        
        Worksheets("Merged DRRF & DTN").Select
        
        If DTNsite <> preDTNsite Then
            t = ActiveSheet.UsedRange.Rows.Count
            'Finds section of same store numbers
            With Worksheets("Merged DRRF & DTN").Range("A1", "a" & t)
                Set s = .Find(what:=DTNsite, LookIn:=xlValues, LookAt:=xlWhole)
                If Not s Is Nothing Then
                    Set sList = s
                    firstAddress = s.Address
                    Do
                        Set sList = Union(sList, s)
                        Worksheets("Merged DRRF & DTN").Range(s.Address).Activate
                        Set s = .FindNext(s)
                    Loop While Not s Is Nothing And s.Address <> firstAddress
                End If
            End With
        End If
        Range(sList, sList.Offset(0, 4)).Select
        
        'Within same store number selection, Finds matching amount
        Dim a As Range
        With Selection
            Set a = .Find(what:=DTNamt, LookIn:=xlFormulas)
            If Not a Is Nothing Then
                DRRFbatch = Range(a.Address).Offset(0, -3)
                If DRRFbatch = "EMS" Or DRRFbatch = DTNbatch Then
                    Range(a.Address).EntireRow.ClearContents
                    Worksheets("DTN").Select
                    Cells(r, 1).EntireRow.ClearContents
                    
                ElseIf DRRFbatch = "P97" Then
                    Range(a.Address).Offset(0, 1).Value = "Matched"
                    Worksheets("DTN").Cells(r, 1).EntireRow.Font.color = vbRed
                End If
            End If
        End With
        
        r = r + 1
        preDTNsite = DTNsite
        Worksheets("DTN").Select
        Cells(r, 1).Select
    
    Loop
    
    
    ' Deletes all blank rows and labels all as DTN
    
    Worksheets("DTN").Select
    t = ActiveSheet.UsedRange.Rows.Count
    Range("a2", "a" & t).Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete (xlUp)
    t = ActiveSheet.UsedRange.Rows.Count
    Cells(1, 4).Select
    Selection.Value = "DTN"
    Selection.Copy
    Range(Cells(2, 4), Cells(t, 4)).PasteSpecial xlPasteValues
    
    
    ' Deletes all blank rows and labels all as DR/FNF
    
    Worksheets("Merged DRRF & DTN").Select
    t = ActiveSheet.UsedRange.Rows.Count
    Range("a2", "a" & t).Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete (xlUp)
    t = ActiveSheet.UsedRange.Rows.Count
    Cells(1, 4).Select
    Selection.Value = "DR/FNF"
    Selection.Copy
    Range(Cells(2, 4), Cells(t, 4)).PasteSpecial xlPasteValues
    
    
    'Formats and copies DTN to DRRF
    
    Range(Cells(1, 1), Cells(1, 5)).ClearFormats
    Range(Cells(1, 5), Cells(t, 5)).ClearFormats
    Cells(1, 1).Value = "Site"
    Cells(1, 2).Value = "Batch"
    Cells(1, 3).Value = "Date"
    Cells(1, 4).Value = "Source"
    Cells(1, 5).Value = "Amt"
    
    
    ' Changes remaining DTN's Dates and copies all over to Merge
    
    Worksheets("DTN").Select
    t = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To t
        If Cells(i, 8).Value <> "" Then
            Cells(i, 3).Value = Cells(i, 8).Value
        End If
    Next i
    Range(Cells(2, 1), Cells(t, 5)).Copy
    
    Worksheets("Merged DRRF & DTN").Select
    t = ActiveSheet.UsedRange.Rows.Count
    Cells(t + 1, 1).PasteSpecial xlPasteAll
    
'    oldvalue = Application.DisplayAlerts
'    Application.DisplayAlerts = False
'    Sheets("DTN").Delete
'    Application.DisplayAlerts = oldvalue
    
    
    ' Changes all with DR/FNF font color to red
    
    r = 2
    Cells(r, 4).Select
    Do While ActiveCell.Value = "DR/FNF"
        ActiveCell.EntireRow.Font.color = vbRed
        r = r + 1
        Cells(r, 4).Select
    Loop
    
    t = ActiveSheet.UsedRange.Rows.Count
    Range("a2:a" & t).HorizontalAlignment = xlRight
    Range("b2:b" & t).HorizontalAlignment = xlRight
    Range("c2:c" & t).HorizontalAlignment = xlRight
    Range("d2:d" & t).HorizontalAlignment = xlCenter
    Range("e2:e" & t).HorizontalAlignment = xlLeft
    Range("e2:e" & t).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    
    ' Adds two columns
    
    Columns("A:B").Insert xlToRight
    Cells(1, 1).Value = "Type"
    Cells(1, 2).Value = "Notes"
    
    'Saves File
    Dim StartDate As String
    Dim EndDate As String
    Dim FileName_Chev_DRSA As String
    Dim FilePath_Chev_DRSA As String
    
    file = ActiveWorkbook.Name
    FileNameParts = Split(file, " ")
    StartDate = FileNameParts(4)
    EndArray = Split(FileNameParts(6), ".")
    EndDate = EndArray(0) & "." & EndArray(1) & "." & EndArray(2)
    
    FileName_Chev_DRSA = "DRRF_DTN Merged " & StartDate & " to " & EndDate
    FilePath_Chev_DRSA = "\\Server\f\Accounting\S.Mitchell\Chevron\DRRF_DRSA_DTN\" & FileName_Chev_DRSA & ".xlsx"
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:=FilePath_Chev_DRSA, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
End Sub

Sub SeparateP97s()

    
    


End Sub

Sub Chevron_Merge_2()

    'Saves File
    Dim StartDate As String
    Dim EndDate As String
    Dim FileName_Chev_DRSA As String
    Dim FilePath_Chev_DRSA As String
    Dim FolderPath As String
    
    file = ActiveWorkbook.Name
    FileNameParts = Split(file, " ")
    StartDate = FileNameParts(4)
    EndArray = Split(FileNameParts(6), ".")
    EndDate = EndArray(0) & "." & EndArray(1) & "." & EndArray(2) ' Avoids including ".xlsx"
    
    FolderPath = "F:\Accounting\S.Mitchell\Chevron\New Merged\CreditCards " & StartDate & " to " & EndDate
    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir (FolderPath)
    End If
    
    FileName_Chev_DRSA = "DR_DTN Totals " & StartDate & " to " & EndDate & ".xlsx"
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:=FolderPath & FileName_Chev_DRSA, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    
    ' -----------------------------------
    ' - Merging DR FNF REG with DTN REG -
    ' -----------------------------------
    
    Worksheets("DR FNF REG").Copy before:=Worksheets("DR FNF REG")
    ActiveSheet.Name = "Merged DRRF & DTN"
    Worksheets("DTN REG").Copy before:=Worksheets("DTN REG")
    ActiveSheet.Name = "Merging DTN REG"
    
    Worksheets("Merging DTN REG").Select
    r = 2
    'Cells(r, 1).Select
    Do While Cells(r, 1).Value <> ""
        
        DTNsite = Cells(r, 1).Value
        DTNbatch = Cells(r, 2).Value
        DTNamt = (Cells(r, 5).Value) * -1
        
        Worksheets("Merged DRRF & DTN").Select
        
        If DTNsite <> preDTNsite Then
            t = ActiveSheet.UsedRange.Rows.Count
            'Finds section of same store numbers
            With Worksheets("Merged DRRF & DTN").Range("A1", "a" & t)
                Set s = .Find(what:=DTNsite, LookIn:=xlValues, LookAt:=xlWhole)
                If Not s Is Nothing Then
                    Set sList = s
                    firstAddress = s.Address
                    Do
                        Set sList = Union(sList, s)
                        Worksheets("Merged DRRF & DTN").Range(s.Address).Activate
                        Set s = .FindNext(s)
                    Loop While Not s Is Nothing And s.Address <> firstAddress
                End If
            End With
        End If
        Range(sList, sList.Offset(0, 4)).Select
        
        'Within same store number selection, Finds matching amount
        Dim a As Range
        With Selection
            Set a = .Find(what:=DTNamt, LookIn:=xlFormulas)
            If Not a Is Nothing Then
                DRRFbatch = Range(a.Address).Offset(0, -3)
                If DRRFbatch = "EMS" Or DRRFbatch = DTNbatch Then
                    Range(a.Address).EntireRow.ClearContents
                    Worksheets("Merging DTN REG").Select
                    Cells(r, 1).EntireRow.ClearContents
                End If
            End If
        End With
        
        r = r + 1
        preDTNsite = DTNsite
        Worksheets("Merging DTN REG").Select
        Cells(r, 1).Select
    
    Loop
    
    
    ' Deletes all blank rows and labels all as DTN
    
    Worksheets("Merging DTN REG").Select
    t = ActiveSheet.UsedRange.Rows.Count
    Range("a2", "a" & t).Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete (xlUp)
    t = ActiveSheet.UsedRange.Rows.Count
    Cells(1, 4).Select
    Selection.Value = "DTN"
    Selection.Copy
    Range(Cells(2, 4), Cells(t, 4)).PasteSpecial xlPasteValues
    
    
    ' Deletes all blank rows and labels all as DR/FNF
    
    Worksheets("Merged DRRF & DTN").Select
    t = ActiveSheet.UsedRange.Rows.Count
    Range("a2", "a" & t).Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete (xlUp)
    t = ActiveSheet.UsedRange.Rows.Count
    Cells(1, 4).Select
    Selection.Value = "DR/FNF"
    Selection.Copy
    Range(Cells(2, 4), Cells(t, 4)).PasteSpecial xlPasteValues
    
    
    'Formats and copies DTN to DRRF
    
    Range(Cells(1, 1), Cells(1, 5)).ClearFormats
    Range(Cells(1, 5), Cells(t, 5)).ClearFormats
    Cells(1, 1).Value = "Site"
    Cells(1, 2).Value = "Batch"
    Cells(1, 3).Value = "Date"
    Cells(1, 4).Value = "Source"
    Cells(1, 5).Value = "Amt"
    
    
    ' Changes remaining DTN's Dates and copies all over to Merge
    
    Worksheets("Merging DTN REG").Select
    t = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To t
        If Cells(i, 8).Value <> "" Then
            Cells(i, 3).Value = Cells(i, 8).Value
        End If
    Next i
    Range(Cells(2, 1), Cells(t, 5)).Copy
    
    Worksheets("Merged DRRF & DTN").Select
    t = ActiveSheet.UsedRange.Rows.Count
    Cells(t + 1, 1).PasteSpecial xlPasteAll
    
    oldvalue = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Sheets("Merging DTN REG").Delete
    Application.DisplayAlerts = oldvalue
    
    ' Changes all with DR/FNF font color to red
    r = 2
    Cells(r, 4).Select
    Do While ActiveCell.Value = "DR/FNF"
        ActiveCell.EntireRow.Font.color = vbRed
        r = r + 1
        Cells(r, 4).Select
    Loop
    
    t = ActiveSheet.UsedRange.Rows.Count
    Range("a2:a" & t).HorizontalAlignment = xlRight
    Range("b2:b" & t).HorizontalAlignment = xlRight
    Range("c2:c" & t).HorizontalAlignment = xlRight
    Range("d2:d" & t).HorizontalAlignment = xlCenter
    Range("e2:e" & t).HorizontalAlignment = xlLeft
    Range("e2:e" & t).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    
    ' Adds two columns
    Columns("A:B").Insert xlToRight
    Cells(1, 1).Value = "Type"
    Cells(1, 2).Value = "Notes"
    
    
    
    
    
    ' -----------------------------------
    ' - Merging DR FNF P97 with DTN P97 -
    ' -----------------------------------
    
    
    Worksheets("DR FNF P97").Copy before:=Worksheets("DR FNF P97")
    ActiveSheet.Name = "Merged P97 & DTN"
    Worksheets("DTN P97").Copy before:=Worksheets("DTN P97")
    ActiveSheet.Name = "Merging DTN P97"
    
    Worksheets("Merging DTN P97").Select
    r = 2
    'Cells(r, 1).Select
    Do While Cells(r, 1).Value <> ""
        
        DTNsite = Cells(r, 1).Value
        DTNbatch = Cells(r, 2).Value
        DTNamt = (Cells(r, 5).Value) * -1
        
        Worksheets("Merged P97 & DTN").Select
        
        If DTNsite <> preDTNsite Then
            t = ActiveSheet.UsedRange.Rows.Count
            'Finds section of same store numbers
            With Worksheets("Merged P97 & DTN").Range("A1", "a" & t)
                Set s = .Find(what:=DTNsite, LookIn:=xlValues, LookAt:=xlWhole)
                If Not s Is Nothing Then
                    Set sList = s
                    firstAddress = s.Address
                    Do
                        Set sList = Union(sList, s)
                        Worksheets("Merged P97 & DTN").Range(s.Address).Activate
                        Set s = .FindNext(s)
                    Loop While Not s Is Nothing And s.Address <> firstAddress
                End If
            End With
        End If
        Range(sList, sList.Offset(0, 4)).Select
        
        'Within same store number selection, Finds matching amount
        'Dim a As Range
        With Selection
            Set a = .Find(what:=DTNamt, LookIn:=xlFormulas)
            If Not a Is Nothing Then
                Range(a.Address).Offset(0, 1).Value = "Matched"
                Worksheets("Merging DTN P97").Select
                Cells(r, 1).EntireRow.Font.color = vbRed
            End If
        End With
        
        r = r + 1
        preDTNsite = DTNsite
        Worksheets("Merging DTN P97").Select
        Cells(r, 1).Select
    
    Loop
    
    
    ' Deletes all blank rows and labels all as DTN
    
    Worksheets("Merging DTN P97").Select
    t = ActiveSheet.UsedRange.Rows.Count
    Range("a2", "a" & t).Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete (xlUp)
    t = ActiveSheet.UsedRange.Rows.Count
    Cells(1, 4).Select
    Selection.Value = "DTN"
    Selection.Copy
    Range(Cells(2, 4), Cells(t, 4)).PasteSpecial xlPasteValues
    
    
    ' Deletes all blank rows and labels all as DR/FNF
    
    Worksheets("Merged P97 & DTN").Select
    t = ActiveSheet.UsedRange.Rows.Count
    Range("a2", "a" & t).Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete (xlUp)
    t = ActiveSheet.UsedRange.Rows.Count
    Cells(1, 4).Select
    Selection.Value = "DR/FNF"
    Selection.Copy
    Range(Cells(2, 4), Cells(t, 4)).PasteSpecial xlPasteValues
    
    
    'Formats and copies DTN to DRRF
    
    Range(Cells(1, 1), Cells(1, 5)).ClearFormats
    Range(Cells(1, 5), Cells(t, 5)).ClearFormats
    Cells(1, 1).Value = "Site"
    Cells(1, 2).Value = "Batch"
    Cells(1, 3).Value = "Date"
    Cells(1, 4).Value = "Source"
    Cells(1, 5).Value = "Amt"
    
    
    ' Changes remaining DTN's Dates and copies all over to Merge
    
    Worksheets("Merging DTN P97").Select
    t = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To t
        If Cells(i, 8).Value <> "" Then
            Cells(i, 3).Value = Cells(i, 8).Value
        End If
    Next i
    Range(Cells(2, 1), Cells(t, 5)).Copy
    
    Worksheets("Merged P97 & DTN").Select
    t = ActiveSheet.UsedRange.Rows.Count
    Cells(t + 1, 1).PasteSpecial xlPasteAll
    
    oldvalue = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Sheets("Merging DTN P97").Delete
    Application.DisplayAlerts = oldvalue
    
    ' Changes all with DR/FNF font color to red
    r = 2
    Cells(r, 4).Select
    Do While ActiveCell.Value = "DR/FNF"
        ActiveCell.EntireRow.Font.color = vbRed
        r = r + 1
        Cells(r, 4).Select
    Loop
    
    t = ActiveSheet.UsedRange.Rows.Count
    Range("a2:a" & t).HorizontalAlignment = xlRight
    Range("b2:b" & t).HorizontalAlignment = xlRight
    Range("c2:c" & t).HorizontalAlignment = xlRight
    Range("d2:d" & t).HorizontalAlignment = xlCenter
    Range("e2:e" & t).HorizontalAlignment = xlLeft
    Range("e2:e" & t).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    
    ' Adds two columns
    Columns("A:B").Insert xlToRight
    Cells(1, 1).Value = "Type"
    Cells(1, 2).Value = "Notes"
    
    
    'Saves File
    ActiveWorkbook.Save
    
    
'Copies REG (Merge, DR, DTN) worksheets and creates new workbook
    
    Dim wbREG As Workbook
    Dim REGfilename As String
    'Dim REGfilepath As String
    
    Dim REG_wsMerge As Worksheet
    Dim REG_wsDR As Worksheet
    Dim REG_wsDTN As Worksheet
    Dim GiftCards As Worksheet
    
    Set REG_wsMerge = Workbooks(originalfilename).Sheets("Merged DRRF & DTN")
    Set REG_wsDR = Workbooks(originalfilename).Sheets("DR FNF REG")
    Set REG_wsDTN = Workbooks(originalfilename).Sheets("DTN REG")
    Set GiftCards = Workbooks(originalfilename).Sheets("Giftcards")
    
    REG_wsMerge.Move
    Set wbREG = ActiveWorkbook
    REG_wsDTN.Move after:=wbREG.Sheets(1)
    REG_wsDR.Move after:=wbREG.Sheets(1)
    GiftCards.Move after:=wbREG.Sheets(wbREG.Sheets.Count)
    
    REGfilename = "DRRF_DTN Merged " & StartDate & " to " & EndDate & ".xlsx"
    Application.DisplayAlerts = False
    wbREG.SaveAs FolderPath & REGfilename
    Application.DisplayAlerts = True
    
    
'Copies P97 (Merge, DR, DTN) worksheets and creates new workbook
    
    Dim wbP97 As Workbook
    Dim P97filename As String
    'Dim P97filepath As String
    
    Dim P97_wsMerge As Worksheet
    Dim P97_wsDR As Worksheet
    Dim P97_wsDTN As Worksheet
    
    Set P97_wsMerge = Workbooks(originalfilename).Sheets("Merged P97 & DTN")
    Set P97_wsDR = Workbooks(originalfilename).Sheets("PDF FNF P97")
    Set P97_wsDTN = Workbooks(originalfilename).Sheets("DTN P97")
    
    P97_wsMerge.Move
    Set wbP97 = ActiveWorkbook
    P97_wsDTN.Move after:=wbP97.Sheets(1)
    P97_wsDR.Move after:=wbP97.Sheets(1)
    
    P97filename = "P97_DTN Merged " & StartDate & " to " & EndDate & ".xlsx"
    Application.DisplayAlerts = False
    wbP97.SaveAs FolderPath & P97filename
    Application.DisplayAlerts = True
    
End Sub


Sub Chev_Saves()
    
    Dim originalfilename As String
    originalfilename = ActiveWorkbook.Name
    
    Dim StartDate As String
    Dim EndDate As String
    Dim FolderPath As String
    
    'Gets Dates
    file = ActiveWorkbook.Name
    FileNameParts = Split(file, " ")
    StartDate = FileNameParts(2)
    EndArray = Split(FileNameParts(4), ".")
    EndDate = EndArray(0) & "." & EndArray(1) & "." & EndArray(2) ' Avoids including ".xlsx"
    
    '   Checks if folder exists and creates it if not
    FolderPath = "F:\Accounting\S.Mitchell\Chevron\New Merged\CreditCards " & StartDate & " to " & EndDate
    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir (FolderPath)
    End If
    
'Copies REG (Merge, DR, DTN) worksheets and creates new workbook
    
    Dim wbREG As Workbook
    Dim REGfilename As String
    'Dim REGfilepath As String
    
    Dim REG_wsMerge As Worksheet
    Dim REG_wsDR As Worksheet
    Dim REG_wsDTN As Worksheet
    Dim GiftCards As Worksheet
    
    Set REG_wsMerge = Workbooks(originalfilename).Sheets("Merged DRRF & DTN")
    Set REG_wsDR = Workbooks(originalfilename).Sheets("DR FNF REG")
    Set REG_wsDTN = Workbooks(originalfilename).Sheets("DTN REG")
    Set GiftCards = Workbooks(originalfilename).Sheets("Giftcards")
    
    REG_wsMerge.Move
    Set wbREG = ActiveWorkbook
    REG_wsDTN.Move after:=wbREG.Sheets(1)
    REG_wsDR.Move after:=wbREG.Sheets(1)
    GiftCards.Move after:=wbREG.Sheets(wbREG.Sheets.Count)
    
    REGfilename = "DRRF_DTN Merged " & StartDate & " to " & EndDate & ".xlsx"
    Application.DisplayAlerts = False
    wbREG.SaveAs FolderPath & REGfilename
    Application.DisplayAlerts = True
    
    
'Copies P97 (Merge, DR, DTN) worksheets and creates new workbook
    
    Dim wbP97 As Workbook
    Dim P97filename As String
    'Dim P97filepath As String
    
    Dim P97_wsMerge As Worksheet
    Dim P97_wsDR As Worksheet
    Dim P97_wsDTN As Worksheet
    
    Set P97_wsMerge = Workbooks(originalfilename).Sheets("Merged P97 & DTN")
    Set P97_wsDR = Workbooks(originalfilename).Sheets("DR FNF P97")
    Set P97_wsDTN = Workbooks(originalfilename).Sheets("DTN P97")
    
    P97_wsMerge.Move
    Set wbP97 = ActiveWorkbook
    P97_wsDTN.Move after:=wbP97.Sheets(1)
    P97_wsDR.Move after:=wbP97.Sheets(1)
    
    P97filename = "P97_DTN Merged " & StartDate & " to " & EndDate & ".xlsx"
    Application.DisplayAlerts = False
    wbP97.SaveAs FolderPath & P97filename
    Application.DisplayAlerts = True


End Sub

Sub P97_PDF_Reports()

' Clean P97 Report:
    ActiveSheet.Copy before:=ActiveSheet
    ActiveSheet.Name = "P97s Transactions & Totals"
'    Worksheets.Add(after:=Worksheets("P97s Transactions")).Name = "Summary Totals"
'    Worksheets("Summary Totals").Select
'    Cells(1, 1).Value = "Transaction Summary Totals"
'    Worksheets("P97s Transactions").Select
        
    ' Edit first column to display sites
    
    Dim Total_Rows As Integer
    Dim Total_Cols As Integer
    
    Total_Rows = Worksheets("P97s Transactions").UsedRange.Rows.Count
    Total_Cols = Worksheets("P97s Transactions").UsedRange.Columns.Count
    
    Dim i As Integer
    
    For i = 2 To Total_Rows
    
        date_entry = Replace(Cells(i, 1).Value, ".", "/", 1)
        Cells(i, 1).Value = date_entry
        
        file = Cells(i, 2).Value
        site = Split(file, ".")
        Cells(i, 2).Value = site(0)
        Cells(i, 2).HorizontalAlignment = xlHAlignLeft
        
    Next i
    
    ' Cleaning
    Dim c As Integer
    Dim j As Byte
    Dim total_amt As String
    
    Dim old_site As String
    Dim new_site As String
    Dim next_site As String
    
    For i = 2 To Total_Rows
    
        old_site = Cells(i - 1, 2).Value
        new_site = Cells(i, 2).Value
        next_site = Cells(i + 1, 2).Value
        
        ' Removes Report Heading
        If Cells(i, 3).Value Like "Merchant ID*" Then
        
            c = i
            Do Until Cells(c, 3).Value Like "Date/Time*"
                c = c + 1
            Loop
            c = c - 1
            Rows(i & ":" & c).ClearContents
            i = c
            
            
        ' Removes all unnessary columns from report
        ElseIf Cells(i, 3).Value Like "Date/Time*" Then
        
            ' Section Rows
            c = i
            Do Until Cells(c, 3).Value Like "Transaction Summary Totals"
                c = c + 1
                If Cells(c, 3).Value Like "Merchant*" Then
                    c = c - 1
                    Exit Do
                End If
            Loop
            c = c - 1
            ' Section Columns
            j = 3
            Do Until Cells(i, j).Value Like "Transaction*" Or Cells(i, j).Value Like "Total*"
                j = j + 1
            Loop
            j = j - 1
            Range(Cells(i, 3), Cells(c, j)).Delete shift:=xlToLeft
            i = c
        
        
        ' Captures Total Amt from Report
        ElseIf Cells(i, 3).Value Like "Transaction Summary Totals" Then
            
            c = i
            Do Until Cells(c, 3).Value Like "Total"
                c = c + 1
                If Cells(c, 3).Value Like "Merchant*" Then
                    Exit Do
                End If
            Loop
            
            If Cells(c, 3).Value Like "Total" Then
                j = 3
                blank = 0
                Do Until Cells(i + 1, j).Value Like "Sales Amount"
                    j = j + 1
                    If Cells(i + 1, j).Value = "" Then
                        blank = blank + 1
                        If blank = Total_Cols Then
                            m = MsgBox("Error in search for Sales Amount. Row: " & i + 1)
                            Exit Do
                        End If
                    End If
                Loop
                
                total_amt = Cells(c, j).Value
                Range(Cells(i, 3), Cells(c, j)).Delete shift:=xlToLeft
                Cells(i, 3).Value = total_amt
                Cells(i, 4).Value = "Total"
                i = c
            End If
            
            
        ' Identifies site with data not imported
        ElseIf new_site <> old_site And new_site <> next_site Then
            Cells(i, 3).Value = "Data Not Imported"
            
            
        ' Deletes the row containing the page count
        ElseIf new_site <> next_site Then
            For Each cell In Range(Cells(i, 1), Cells(i, Total_Cols)).Cells
                If cell.Value Like "Page*" Then
                    If cell.Value Like "*2" Then
                        Range(Cells(i, 3), Cells(i, Total_Cols)).ClearContents
                        Cells(i, 3).Value = "Missing Data"
                        Exit For
                    Else
                        Range(Cells(i, 3), Cells(i, Total_Cols)).ClearContents
                        Exit For
                    End If
                End If
            Next cell
        End If
    Next i
    
    
    For i = 2 To Total_Rows
        If Cells(i, 3).Value Like "Transaction*" Then
            Rows(i).ClearContents
        ElseIf Cells(i, 3).Value Like "Total*" Then
            Rows(i).ClearContents
        Else
            ' Keep rows with amounts
        End If
    Next i
    
    ' Formats neccessary columns
    Columns(3).SpecialCells(xlCellTypeBlanks).EntireRow.Delete (xlUp)
    Columns(3).HorizontalAlignment = xlHAlignRight
    
    Range(Cells(1, 5), Cells(1, Total_Cols)).EntireColumn.Delete
    Range(Cells(1, 1), Cells(1, 4)).EntireColumn.Insert
    Columns(6).Copy
    Columns(3).PasteSpecial xlPasteValues
    Columns(6).ClearContents
    
    Cells(1, 1).Value = "Type"
    Cells(1, 2).Value = "Notes"
    Cells(1, 3).Value = "Site"
    Cells(1, 4).Value = "Batch"
    Cells(1, 5).Value = "Date"
    Cells(1, 6).Value = "Source"
    Cells(1, 7).Value = "Amount"
    
    Columns(1).ColumnWidth = 15
    Columns(2).ColumnWidth = 15
    Columns(3).ColumnWidth = 15
    Columns(4).ColumnWidth = 25
    Columns(5).ColumnWidth = 15
    Columns(6).ColumnWidth = 15
    Columns(7).ColumnWidth = 20
    
    Dim amt As Variant
    
    Total_Rows = Worksheets("P97s Transactions").UsedRange.Rows.Count
    Total_Cols = Worksheets("P97s Transactions").UsedRange.Columns.Count
    
    For i = 2 To Total_Rows
        
        ' Populates Batch Column
        If Cells(i, 8).Value Like "Total" Then
            Cells(i, 4).Value = "P97 Total"
        ElseIf IsNumeric(Cells(i, 7).Value) Then
            Cells(i, 4).Value = "P97 Transaction"
        Else
            ' Do Nothing
        End If
        
        ' Populates Source Column
        Cells(i, 6).Value = "PDF/FNF"
        
        ' Changes amt to negative
        amt = Cells(i, 7).Value
        If IsNumeric(amt) Then
            Cells(i, 7).Value = amt * -1
            Cells(i, 7).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Cells(i, 7).EntireRow.Font.color = vbRed
        End If
    Next i
End Sub








































