Attribute VB_Name = "Report"
Sub generateReport()
    Dim createTable As Boolean, repRange As Range
    Call clearExpenseConsolidate
    Call clearPivotTables
    
    createTable = extractAllExpenseData()
    
    If (createTable = True) Then
        Dim table As PivotTable
        Set table = createExpenseReportPivotTable()
    End If
        
    Set repRange = obterRangeTabela(Sheets(Defs.SHEET_REPORT), Defs.REPORT_EXPENSE_START_LINE, Defs.REPORT_EXPENSE_START_COL, 13)
    If (Not repRange Is Nothing) Then
        repRange.Delete shift:=xlUp
    End If

    If (createTable = True) Then
        Call copyTableDataToReport(table)
    End If

    Call clearExpenseConsolidate
    Call clearPivotTables
End Sub


Private Function extractAllExpenseData() As Boolean
    Dim sheetNames() As Variant, curName As Variant, createTable As Boolean
    Dim curSheet As Worksheet, consSheet As Worksheet
    createTable = False
    sheetNames = getMonthSheets()
    Set consSheet = Worksheets(Defs.SHEET_CONSOLIDATE)
    Call clearExpenseConsolidate
    
    For Each curName In sheetNames
        Dim rgExpense As Range, regMonthList() As MonthRecord
        
        Set curSheet = Worksheets(curName)
        Set rgExpense = obterRangeDespesas(curSheet)
        If (rgExpense Is Nothing) Then
            Dim regMonth As MonthRecord, reg As RegistroValor
            ReDim regMonthList(1)
            
            reg.Categoria = ""
            reg.type = EXPENSE
            reg.valor = 0
            regMonth.record = reg
            regMonth.month = curName
            regMonthList(0) = regMonth
            
        Else
            regMonthList = extractDataFromRange(rgExpense, curSheet.name)
            createTable = True
        End If
        
        Call insertConsolidate(regMonthList)
    Next
    extractAllExpenseData = createTable
End Function

Private Function extractDataFromRange(rg As Range, sheetName As String) As MonthRecord()
    ReDim regMonthList(rg.Rows.Count) As MonthRecord
    
    For i = 1 To rg.Rows.Count
        Dim reg As RegistroValor, regMonth As MonthRecord, rw As Range
        Set rw = rg.Rows(i)
    
        With rw
            reg.diaDoMes = .Cells(2)
            reg.membro = .Cells(3)
            reg.Categoria = .Cells(4)
            reg.descricao = .Cells(5)
            reg.valor = .Cells(6)
            reg.type = EXPENSE
        End With
    
        regMonth.month = sheetName
        regMonth.record = reg
        regMonthList(i - 1) = regMonth
    Next
    extractDataFromRange = regMonthList
End Function

Private Sub insertConsolidate(regList() As MonthRecord)
    Dim firstRow As Integer, lastRow As Integer, startCol As Integer, conSheet As Worksheet
    
    Set conSheet = Worksheets(Defs.SHEET_CONSOLIDATE)
    startCol = Defs.CONSOLIDATE_EXPENSE_START_COL
    firstRow = obterLinhaFinalTabela(conSheet, Defs.CONSOLIDATE_EXPENSE_START_LINE, startCol) + 1
    lastRow = firstRow + UBound(regList) - LBound(regList) - 1
                                    
    For i = firstRow To lastRow
        Dim rec As MonthRecord, rgRow As Range
        rec = regList(i - firstRow)
        Set rgRow = conSheet.Range(conSheet.Cells(i, startCol), conSheet.Cells(i, startCol + 3))
        
        With rgRow
            .Cells(1).value = rec.month
            .Cells(2).value = rec.record.Categoria
            .Cells(3).value = rec.record.membro
            .Cells(4).value = rec.record.valor
        End With
    Next
End Sub

Private Sub clearExpenseConsolidate()
    Dim mSheet As Worksheet
    Set mSheet = Worksheets(Defs.SHEET_CONSOLIDATE)
    mSheet.Range( _
        mSheet.Cells(Defs.CONSOLIDATE_EXPENSE_START_LINE, Defs.CONSOLIDATE_EXPENSE_START_COL), _
        mSheet.Cells(SHEET_ROWCOUNT, Defs.CONSOLIDATE_EXPENSE_START_COL + 3) _
    ).Clear
End Sub

Private Sub clearPivotTables()
    Dim table As PivotTable
    For Each table In Sheets(Defs.SHEET_CONSOLIDATE).PivotTables
        table.TableRange2.Clear
    Next table

End Sub


Private Sub copyTableDataToReport(table As PivotTable)
    Dim tableRange As Range, dataRange As Range, rowCount As Integer, colCount As Integer
    Set tableRange = table.TableRange2
    rowCount = tableRange.Rows.Count
    colCount = tableRange.Columns.Count
    addr = Range(Cells(3, 1), Cells(rowCount, colCount)).Address(rowAbsolute:=False, columnAbsolute:=False)
    Set dataRange = tableRange.Range(addr)
    dataRange.Copy destination:=Sheets(Defs.SHEET_REPORT).Cells(Defs.REPORT_EXPENSE_START_LINE, Defs.REPORT_EXPENSE_START_COL)
End Sub
