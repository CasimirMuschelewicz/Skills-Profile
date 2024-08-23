Attribute VB_Name = "Lottery"
Sub FormatLottery()

    'Opens all necessary files
    Dim wb As Workbook
    Dim FolderPath As String
    Dim FilePath As String
    FolderPath = "C:\Users\smitchell\Desktop\Outlook Attachments\Lottery\"
    FilePath = Dir(FolderPath & "InvoiceSummary*" & ".csv")
    Do While FilePath <> ""
        Set wb = Workbooks.Open(FolderPath & FilePath)
        FilePath = Dir
    Loop
    ActiveSheet.Name = "Invoice Summary"
    InvSumWbName = ActiveWorkbook.Name
    
    FolderPath = "C:\Users\smitchell\Desktop\Outlook Attachments\Lottery\"
    FilePath = Dir(FolderPath & "InvoiceDetail*" & ".csv")
    Do While FilePath <> ""
        Set wb = Workbooks.Open(FolderPath & FilePath)
        FilePath = Dir
    Loop
    ActiveSheet.Name = "Invoice Detail"
    ActiveWorkbook.Sheets("Invoice Detail").Move after:=Workbooks(InvSumWbName).Sheets("Invoice Summary")
    
    FolderPath = "C:\Users\smitchell\Desktop\Outlook Attachments\Lottery\"
    FilePath = Dir(FolderPath & "RetailerPackInventory*" & ".csv")
    Do While FilePath <> ""
        Set wb = Workbooks.Open(FolderPath & FilePath)
        FilePath = Dir
    Loop
    ActiveSheet.Name = "Retailer Pack Inventory"
    ActiveWorkbook.Sheets("Retailer Pack Inventory").Move after:=Workbooks(InvSumWbName).Sheets("Invoice Detail")

    Dim EndDate As String

    file = ActiveWorkbook.Name
    FileName = Split(file, ".")
    FileNameParts = Split(FileName(0), "_")
    EndDate = FileNameParts(2)
    
    endMM = Mid(EndDate, 3, 2)
    endDD = Mid(EndDate, 5, 2)
    endYY = Mid(EndDate, 1, 2)
    
    EndDate = endMM & "." & endDD & "." & endYY
    
    
    
    
    
    
    
    '----------------
    '-Invoice Detail-
    '----------------
    
    ActiveWorkbook.Sheets("Invoice Detail").Select
    
    Range("A:A,B:B,D:D").Select
    Selection.Delete shift:=xlToLeft
    
    Range("D:D").Select
    Selection.Insert shift:=xlToRight
    Range("D1").FormulaR1C1 = "TOTAL ONLINE SALES"
    
    totrow = ActiveSheet.UsedRange.Rows.Count
    
    Range("D2").Select
    Selection.Formula = "=SUM(E2:O2)"
    Selection.Copy
    Range("D3", "D" & totrow).Select
    ActiveSheet.Paste
    Range("D2", "D" & totrow).Copy
    Range("D2", "D" & totrow).PasteSpecial xlPasteValues
    Range("D2", "D" & totrow).Select
    Selection.NumberFormat = "General"
    Columns("E:O").Delete shift:=xlToRight
    Range("H2", "J" & totrow).Select
    Selection.ClearContents
    Columns("K:M").Delete shift:=xlToRight
    Range("L1").Select
    Selection.Formula = "DATE TOTAL"
    Rows("1:1").RowHeight = 45
    Rows("1:1").Select
    Selection.Replace what:="_", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Columns("A:C").ColumnWidth = 15
    Columns("D:L").ColumnWidth = 10
    
    'Deletes non-active sites (sites with zeros across all columns)
    For r = 2 To totrow
        If Range("D" & r).Value = 0 Then
             Range("D" & r).ClearContents
        Else
            '...
        End If
    Next r
    Range("D2", "D" & totrow).Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete shift:=xlUp
    Range("A1").Select
    
    'Gives total of columns of total row and gives Grand Total
    t = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To t
         Cells(i, "L").Formula = "=SUM(D" & i & ":K" & i & ")"
    Next i
    Range("L2:L" & t).Copy
    Range("L2:L" & t).PasteSpecial xlPasteValues
    Cells(t + 2, "L").Formula = "=SUM(L2:L" & t & ")"
    Cells(t + 2, "L").Copy
    Cells(t + 2, "L").PasteSpecial xlPasteValues
    
    
    
    
    
    
    '-------------------------
    '-Retailer Pack Inventory-
    '-------------------------
    Sheets("Retailer Pack Inventory").Select
    
    'Delete unnecessary columns and format rest
    Pack_totrow = ActiveSheet.UsedRange.Rows.Count
    Cells.ColumnWidth = 15
    Range("A:A,B:B,D:D,F:F,K:K,M:M").Select
    Selection.Delete shift:=xlToLeft
    Rows("1:1").Select
    With Selection
        .Replace what:="_", Replacement:=" ", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("C:C").Select
    Selection.WrapText = True
    Selection.ColumnWidth = 1
    Cells.RowHeight = 15
    Rows("1:1").RowHeight = 45
    
    'Moves Date Settled
    Columns("D:D").Select
    Selection.Insert shift:=xlToRight
    Selection.ColumnWidth = 15
    Columns("J:J").Select
    Selection.Copy
    Columns("D:D").Select
    Selection.PasteSpecial xlPasteValues
    Selection.NumberFormat = "m/d/yyyy"
    Selection.HorizontalAlignment = xlRight
    Columns("J:J").Select
    Selection.Delete shift:=xlToLeft
    
    'A/P column added
    Columns("E:H").Select
    Selection.ColumnWidth = 1
    Columns("I:I").Select
    Selection.Insert shift:=xlToRight
    Selection.ColumnWidth = 15
    Range("I1").Value = "A/P"
    Range("I2").Select
    Selection.Formula = "=E2&""-""&G2"
    Selection.Copy
    Range("I3", "I" & Pack_totrow).Select
    Selection.PasteSpecial xlPasteFormulas
    Range("I2", "I" & Pack_totrow).Select
    Selection.Copy
    Selection.PasteSpecial xlPasteValues
    Selection.HorizontalAlignment = xlRight
        
    'deleted rows without dates settled
    Range("D2", "D" & Pack_totrow).Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete shift:=xlUp
    Pack_totrow = ActiveSheet.UsedRange.Rows.Count
    
    'Sorts by store then by date
    Range("a2", "l" & Pack_totrow).Select
    Selection.Sort Key1:=Range("B2"), Order1:=xlAscending, Header:=xlNo, _
                   Key2:=Range("J2"), Order1:=xlAscending, Header:=xlNo
    
    'Add columns: Trade Cost, Non Trade Cost, Comm
    Range("M1").Value = "TRADE COST"
    Range("N1").Value = "NON TRADE COST"
    Range("O1").Value = "COMM"
    
    'Add formula to "TRADE COST"
    Range("M2").Select
    Selection.Formula = "=L2*0.95"
    Selection.Copy
    Range("M3", "M" & Pack_totrow).Select
    Selection.PasteSpecial xlPasteFormulas
    Range("M2", "M" & Pack_totrow).Copy
    Range("M2", "M" & Pack_totrow).PasteSpecial xlPasteValues
    Cells(Pack_totrow + 2, "M").Formula = "=SUM(M2:M" & Pack_totrow & ")" 'Grand Total
    Cells(Pack_totrow + 2, "M").Copy
    Cells(Pack_totrow + 2, "M").PasteSpecial xlPasteValues
    
    Sheets("Invoice Detail").Select
    Cells(t + 2, "L").Copy
    Sheets("Retailer Pack Inventory").Select
    Cells(Pack_totrow + 3, "M").PasteSpecial xlPasteValues
    Cells(Pack_totrow + 3, "M").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Cells(Pack_totrow + 4, "M").Formula = "=SUM(M" & Pack_totrow + 2 & ",M" & Pack_totrow + 3 & ")"
    Cells(Pack_totrow + 4, "M").Copy
    Cells(Pack_totrow + 4, "M").PasteSpecial xlPasteValues
    
    
    
    
    
    
    
    '---------------------
    '-Checks Debit Totals-
    '---------------------
    Sheets("Invoice Summary").Select
    Columns("C:C").Select
    Dim a As Range
    With Selection
            Set a = .Find(what:="Total Debits", LookIn:=xlValues)
            If Not a Is Nothing Then
                Range(a.Address).Offset(0, 3).Copy
            End If
        End With
    Sheets("Retailer Pack Inventory").Select
    t = ActiveSheet.UsedRange.Rows.Count
    'MsgBox (t)
    Cells(t, "n").PasteSpecial xlPasteValues
    
    'Cells(t, "p").Formula = "=IF(M" & t & "=N" & t & ", ""Equal"", ""Not Equal"")"
    
    'Dim amt1 As Double
    'Dim amt2 As Double
    
    amt1 = Round(Cells(t, "m"), 2)
    amt2 = Round(Cells(t, "n"), 2)
    
    'Dim filename As String
    'Dim FilePath As String
    Dim sourceWs As Worksheet
    Dim csvPath As String
    
    'Saves File for Main File
    Dim mainfilename As String
    mainfilename = "Lottery Detail & Inventory w.e. " & EndDate & ".xlsx"
    FilePath = "\\Server\f\Accounting\S.Mitchell\Lottery\Lottery Tetail & Inventory\" & _
               mainfilename
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:=FilePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    If amt1 = amt2 Then
    
        MsgBox ("Debits Match")
        LottoMatches
    
    Else
        MsgBox ("Debits not Matching")
        'Can create a button that can be pressed when problem is fixed
        
        Exit Sub
        
    End If
    
End Sub

Sub LottoMatches()
    
    Dim EndDate As String
    Dim mainfilename As String
    
    file = ActiveWorkbook.Name
    FileName = Split(file, " ")
    EndDate = FileName(5)
    
    endMM = Mid(EndDate, 1, 2)
    endDD = Mid(EndDate, 4, 2)
    endYY = Mid(EndDate, 7, 2)
    
    EndDate = endMM & "." & endDD & "." & endYY
    
    mainfilename = "Lottery Detail & Inventory w.e. " & EndDate & ".xlsx"

    'Formats DR Audit-------------------------------------------------------
        
    'Copies Invoice Detail and Creates new book for DR Audit
    Dim wsOriginal As Worksheet
    Dim wbAudit As Workbook
    Dim wsCopy As Worksheet
    Dim auditfilename As String
    Set wsOriginal = ActiveWorkbook.Sheets("Invoice Detail")
    wsOriginal.Copy
    Set wbAudit = ActiveWorkbook
    Set wsCopy = wbAudit.Sheets(1)
    wsCopy.Name = "DR Audit"
    auditfilename = "Lottery Audit w.e. " & EndDate & ".xlsx"
    Application.DisplayAlerts = False
    wbAudit.SaveAs "\\Server\f\Accounting\S.Mitchell\Lottery\Devin\" & auditfilename
    Application.DisplayAlerts = True
    
    Columns("F:F").Delete
    Columns("G:K").Delete
    
    t = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To t
        Cells(i, 2).Select
        site = ActiveCell.Value
        site = Replace(site, "Food-N-Fun #", "", 1, 12)
        ActiveCell.Value = Mid(site, 1, 2)
        Selection.HorizontalAlignment = xlCenter
    Next i
    
    
    'Pivot Table
    Dim PTableDate As PivotTable    'Refers to Pivot Table
    Dim PTableTime As PivotTable
    Dim PCache As PivotCache    'Holds the Data the Pivot Table refers to
    Dim PRange As Range         'The range of raw data
    Dim PSheet As Worksheet     'The sheet with the pivot table
    Dim DSheet As Worksheet     'The sheet with the data
    Dim LR As Long              'Last Row
    Dim LC As Long              'Last Column
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Worksheets.Add before:=ActiveSheet ' This will add new worksheet
    ActiveSheet.Name = "TOTALS" ' This will rename the worksheet as "Pivot Sheet"
    On Error GoTo 0
    
    Set PSheet = Worksheets("TOTALS")
    Set DSheet = Worksheets("DR Audit")
    'Find Last used row and column in data sheet
    LR = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
    LC = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    'Set the pivot table data range
    Set PRange = DSheet.Cells(1, 1).Resize(LR, LC)
    'Set pivot cahe
    Set PCache = ActiveWorkbook.PivotCaches.Create(xlDatabase, SourceData:=PRange)
    
    'Create Pivot Table Totals by Date-----------------------------------------
    'Create blank pivot table
    Set PTableDate = PCache.CreatePivotTable(TableDestination:=PSheet.Cells(3, 1), _
                     TableName:="Totals by Site")
    'Insert to Row Filed
    With PSheet.PivotTables("Totals by Site").PivotFields("Name")
    .Orientation = xlRowField
    .Position = 1
    End With
    'Insert Segment to Column Filed & position 1
    With PSheet.PivotTables("Totals by Site").PivotFields("TOTAL ONLINE SALES")
    .Orientation = xlDataField
    .Position = 1
    End With
    With PSheet.PivotTables("Totals by Site").PivotFields("Online Cashes")
    .Orientation = xlDataField
    .Position = 1
    End With
    With PSheet.PivotTables("Totals by Site").PivotFields("Instant Cashes")
    .Orientation = xlDataField
    .Position = 1
    End With
    'Format Pivot Table
    PSheet.PivotTables("Totals by Site").ShowTableStyleRowStripes = False
    PSheet.PivotTables("Totals by Site").TableStyle2 = "PivotStyleMedium9"
    Cells(2, 1).Value = "Totals by Site"
    With Range(Cells(2, 1), Cells(2, 4))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = True
        .Font.Bold = True
        .Font.Size = 16
    End With
    With Range(Cells(4, 1), Cells(4, 4))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    Columns("b:d").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    'Create Pivot Table Totals by Date & Time-------------------------------------
    'Create blank pivot table
    Set PTableDate = PCache.CreatePivotTable(TableDestination:=PSheet.Cells(3, 6), _
                     TableName:="Totals by Site & Date")
    'Insert to Row Filed
    With PSheet.PivotTables("Totals by Site & Date").PivotFields("Name")
    .Orientation = xlRowField
    .Position = 1
    End With
    With PSheet.PivotTables("Totals by Site & Date").PivotFields("Date")
    .Orientation = xlRowField
    .Position = 2
    End With
    'Insert Segment to Column Filed & position 1
    With PSheet.PivotTables("Totals by Site & Date").PivotFields("TOTAL ONLINE SALES")
    .Orientation = xlDataField
    .Position = 1
    End With
    With PSheet.PivotTables("Totals by Site & Date").PivotFields("Online Cashes")
    .Orientation = xlDataField
    .Position = 1
    End With
    With PSheet.PivotTables("Totals by Site & Date").PivotFields("Instant Cashes")
    .Orientation = xlDataField
    .Position = 1
    End With
    'Format Pivot Table
    PSheet.PivotTables("Totals by Site & Date").ShowTableStyleRowStripes = False
    PSheet.PivotTables("Totals by Site & Date").TableStyle2 = "PivotStyleMedium9"
    Cells(2, 6).Value = "Totals by Site & Date"
    With Range(Cells(2, 6), Cells(2, 9))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = True
        .Font.Bold = True
        .Font.Size = 16
    End With
    With Range(Cells(4, 6), Cells(4, 9))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    Columns("g:i").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Columns("A:A").ColumnWidth = 10
    Columns("B:D").ColumnWidth = 15
    Columns("F:F").ColumnWidth = 15
    Columns("G:I").ColumnWidth = 15
    
    Cells(2, 1).EntireRow.RowHeight = 21
    
    ActiveWorkbook.Save
    
    
    
    
    
    
    
    
    
    'Formats APVTI Import---------------------------------------------------------
    Workbooks(mainfilename).Activate
    Sheets("Retailer Pack Inventory").Select
    
    'Copies Retailer Pack Inventory and Creates new book for Import
    Dim APVTIfilename As String
    Set wsOriginal = ActiveWorkbook.Sheets("Retailer Pack Inventory")
    wsOriginal.Copy
    Set wbImport = ActiveWorkbook
    Set wsCopy = wbImport.Sheets(1)
    wsCopy.Name = "Import"
    APVTIfilename = "Lottery Import w.e. " & EndDate & ".xlsx"
    Application.DisplayAlerts = False
    wbImport.SaveAs "\\Server\f\Accounting\S.Mitchell\Lottery\Import to APVTI\" & APVTIfilename
    Application.DisplayAlerts = True
    
    ' Delete columns
    Columns("c:h").Select
    Selection.Delete
    Columns("e:f").Select
    Selection.Delete
    Columns("f:g").Select
    Selection.Delete
    r = ActiveSheet.UsedRange.Rows.Count
    Range("A" & r - 2 & ":A" & r).EntireRow.Select
    Selection.Delete
    
    ' Clean Site Column cells
    t = Range("B" & Rows.Count).End(xlUp).Row
    For i = 2 To t
        Cells(i, 2).Select
        site = ActiveCell.Value
        site = Replace(site, "Food-N-Fun #", "", 1, 12)
        ActiveCell.Value = Mid(site, 1, 2)
        Selection.HorizontalAlignment = xlCenter
    Next i
        
    ' Label Columns
    Cells(1, 1).EntireRow.Select
    Selection.ClearContents
    Cells(1, 1).Value = "EFT Date"
    Cells(1, 2).Value = "Name"
    Cells(1, 3).Value = "Reference No."
    Cells(1, 4).Value = "Invoice Date"
    Cells(1, 5).Value = "Amount"
    Cells(1, 6).Value = "Vendor No."
    
    ' Copy Invoice Detail Total and create last line entry
    Workbooks(mainfilename).Activate
    Sheets("Invoice Detail").Select
    t = ActiveSheet.UsedRange.Rows.Count
    Cells(t, "l").Copy
    Workbooks(APVTIfilename).Activate
    Sheets("Import").Select
    t = Range("A" & Rows.Count).End(xlUp).Row
    Cells(t + 1, "e").PasteSpecial xlPasteValues
    Cells(t + 1, "a").Value = Cells(t, "a").Value ' EFT Date
    Cells(t + 1, "b").Value = 2                   ' Site Num
    Cells(t + 1, "c").Value = endMM & endDD & "20" & endYY ' Ref Num
    Cells(t + 1, "d").Value = Cells(t, "a").Value ' Invoice Date
    
    t = Range("A" & Rows.Count).End(xlUp).Row
    Cells(2, 6).Value = "7"
    Range(Cells(2, 6), Cells(t, 6)).FillDown
    
    Columns("B:B").HorizontalAlignment = xlCenter
    
    Workbooks(APVTIfilename).Save
    
    
    
    
' Formats APII Import---------------------------------------------------------

    
    ' Opens APII import template
    Dim TempWb As Workbook
    FolderPath = "\\Server\f\Accounting\S.Mitchell\Lottery\Import to APII\"
    FileName = "Online Lottery Invoice Template.xltx"
    FilePath = FolderPath & FileName
    Set wbInv = Workbooks.Open(FilePath)
    'Workbooks("Online Lottery Invoice Template*.xlsx").Activate
    
    
    ' Creates new file as APII import
    Sheets(1).Name = "Online Lottery Invoice " & EndDate
    APIIfilename = "Online Lottery Invoice " & EndDate
    Application.DisplayAlerts = False
    wbInv.SaveAs "\\Server\f\Accounting\S.Mitchell\Lottery\Import to APII\" & APIIfilename & ".csv"
    Application.DisplayAlerts = True
    
    
    ' Copies Invoice Detail Total from APVTIfilename over to necessary cell on APIIfilename
    apvti_row = Workbooks(APVTIfilename).Sheets(1).UsedRange.Rows.Count
    apvti_col = Workbooks(APVTIfilename).Sheets(1).Cells.Find(what:="Amount", LookIn:=xlValues).Column
    apii_row = Workbooks(APIIfilename).Sheets(1).Cells.Find(what:="Lottery Invoice Total", LookIn:=xlValues).Row
    apii_col = Workbooks(APIIfilename).Sheets(1).Cells.Find(what:="Lottery Invoice Total", LookIn:=xlValues).Column
    Workbooks(APIIfilename).Sheets(1).Cells(apii_row, apii_col).Value = Workbooks(APVTIfilename).Sheets(1).Cells(apvti_row, apvti_col).Value
    
    
    ' Replaces invoice num and date
    Workbooks(APIIfilename).Activate
    Total_Rows = Sheets(1).UsedRange.Rows.Count
    For Each cell In Range(Cells(1, 3), Cells(Total_Rows, 3)).Cells
        cell.Value = endMM & endDD & "20" & endYY
    Next cell
    
    Cells(1, 6).Value = endMM & "/" & endDD & "/" & "20" & endYY
    Cells(1, 7).Value = endMM & "/" & endDD & "/" & "20" & endYY
    Cells(1, 1).Select
    
    'Saves file
    Workbooks(APIIfilename).Save
    
    
    
    

End Sub


Sub Compare_Multi_Columns()

        unpaid_totrow = Worksheets("Unpaids").UsedRange.Rows.Count
        For i = 2 To unpaid_totrow
        
            unpaid_date = Worksheets("Unpaids").Cells(i, 4).Value
            site = Worksheets("Unpaids").Cells(i, 6).Value
            amt = Worksheets("Unpaids").Cells(i, 8).Value
            lottery_totrow = Worksheets("Lottery").UsedRange.Rows.Count
            
            For j = 2 To lottery_totrow
                
                Lot_date = Worksheets("Lottery").Cells(j, 2).Value
                Lot_Site = Worksheets("Lottery").Cells(j, 3).Value
                Lot_amt = Worksheets("Lottery").Cells(j, 6).Value
                
                If site = Lot_Site And amt = Lot_amt And (unpaid_date - Lot_date <= 7 And unpaid_date - Lot_date >= -7) Then
                    
                    Worksheets("Lottery").Cells(j, 11).Value = Worksheets("Unpaids").Cells(i, 1).Value
                    Worksheets("Lottery").Cells(j, 12).Value = Worksheets("Unpaids").Cells(i, 2).Value
                    Worksheets("Lottery").Cells(j, 13).Value = Worksheets("Unpaids").Cells(i, 3).Value
                    Worksheets("Lottery").Cells(j, 14).Value = Worksheets("Unpaids").Cells(i, 4).Value
                    Worksheets("Lottery").Cells(j, 15).Value = Worksheets("Unpaids").Cells(i, 5).Value
                    Worksheets("Lottery").Cells(j, 16).Value = Worksheets("Unpaids").Cells(i, 6).Value
                    Worksheets("Lottery").Cells(j, 17).Value = Worksheets("Unpaids").Cells(i, 7).Value
                    Worksheets("Lottery").Cells(j, 18).Value = Worksheets("Unpaids").Cells(i, 8).Value
                    Worksheets("Lottery").Cells(j, 19).Value = Worksheets("Unpaids").Cells(i, 9).Value
                    Worksheets("Lottery").Cells(j, 20).Value = Worksheets("Unpaids").Cells(i, 10).Value
                    Worksheets("Lottery").Cells(j, 21).Value = Worksheets("Unpaids").Cells(i, 11).Value
                    
                    Exit For
            
                End If
            
            Next j
            
        Next i
        
End Sub

Sub LotteryUnpaid()
        
        Dim Unpaid_inv As String
        Dim Unpaid_amt As String
        Dim Lot_inv As String
        Dim Lot_amt As String
        
        unpaid_totrow = Worksheets("Unpaid").UsedRange.Rows.Count
        lottery_totrow = Worksheets("Lottery").UsedRange.Rows.Count
        For i = 2 To unpaid_totrow
        
            Unpaid_inv = Worksheets("Unpaid").Cells(i, 3).Value
            Unpaid_amt = Worksheets("Unpaid").Cells(i, 8).Value
            
            For j = 2 To lottery_totrow
                
                Lot_inv = Worksheets("Lottery").Cells(j, 4).Value
                Lot_amt = Worksheets("Lottery").Cells(j, 6).Value
                
                If Unpaid_inv = Lot_inv And Unpaid_amt = Lot_amt Then
                    
                    Worksheets("Lottery").Cells(j, 9).Value = "Matched"
                    
                    Worksheets("Lottery").Cells(j, 11).Value = Worksheets("Unpaid").Cells(i, 1).Value
                    Worksheets("Lottery").Cells(j, 12).Value = Worksheets("Unpaid").Cells(i, 2).Value
                    Worksheets("Lottery").Cells(j, 13).Value = Worksheets("Unpaid").Cells(i, 3).Value
                    Worksheets("Lottery").Cells(j, 14).Value = Worksheets("Unpaid").Cells(i, 4).Value
                    Worksheets("Lottery").Cells(j, 15).Value = Worksheets("Unpaid").Cells(i, 5).Value
                    Worksheets("Lottery").Cells(j, 16).Value = Worksheets("Unpaid").Cells(i, 6).Value
                    Worksheets("Lottery").Cells(j, 17).Value = Worksheets("Unpaid").Cells(i, 7).Value
                    Worksheets("Lottery").Cells(j, 18).Value = Worksheets("Unpaid").Cells(i, 8).Value
                    Worksheets("Lottery").Cells(j, 19).Value = Worksheets("Unpaid").Cells(i, 9).Value
                    Worksheets("Lottery").Cells(j, 20).Value = Worksheets("Unpaid").Cells(i, 10).Value
                    Worksheets("Lottery").Cells(j, 21).Value = Worksheets("Unpaid").Cells(i, 11).Value
                    
                    Worksheets("Unpaid").Cells(i, 12).Value = "Matched - Paid"
                    Worksheets("Unpaid").Cells(i, 13).Value = Worksheets("Lottery").Cells(j, 2).Value
                    
                    Exit For
            
                End If
            
            Next j
            
        Next i
        
        m = MsgBox("Matching Complete", , "Status")
End Sub

Sub LotteryUnpaid_RawRetail()
        
        Dim Unpaid_inv As String
        ' Dim Unpaid_amt As String
        Dim Raw_inv As String
        ' Dim Lot_amt As String
        Dim c As Integer
        
        unpaid_totrow = Worksheets("Unpaid").UsedRange.Rows.Count
        lottery_totrow = Worksheets("Retail Raw Files").UsedRange.Rows.Count
        For i = 2 To unpaid_totrow
        
            Unpaid_inv = Worksheets("Unpaid").Cells(i, 4).Value
            'Unpaid_amt = Worksheets("Unpaid").Cells(i, 8).Value
            
            c = 0
            
            For j = 2 To lottery_totrow
                
                Raw_inv = Worksheets("Retail Raw Files").Cells(j, 10).Value
                'Lot_amt = Worksheets("Lottery").Cells(j, 6).Value
                
                If Unpaid_inv = Raw_inv Then
                    
                    Worksheets("Unpaid").Cells(i, 8).Value = Worksheets("Retail Raw Files").Cells(j, 4).Value
                    If Worksheets("Retail Raw Files").Cells(j, 15).Value = "" Then
                        Worksheets("Unpaid").Cells(i, 9).Value = "Not Settled"
                    Else
                        Worksheets("Unpaid").Cells(i, 11).Value = "Settled"
                    End If
                    
                    c = c + 1
            
                End If
            
            Next j
            If c > 0 Then
                Worksheets("Unpaid").Cells(i, 10).Value = c
            End If
        Next i
        
        m = MsgBox("Matching Complete", , "Status")
End Sub












































