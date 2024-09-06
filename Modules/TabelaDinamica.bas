Attribute VB_Name = "TabelaDinamica"
Sub updateCategoryPivotTable()
    Call updateMonthPivotTable( _
        ActiveSheet, _
        Defs.PIVOT_TABLE_CATEG_NOME, _
        obterRangeDespesasAnalise(ActiveSheet), _
        Defs.PIVOT_TABLE_CATEG_LINHA, _
        Defs.PIVOT_TABLE_CATEG_COLUNA, _
        "Categoria", _
        "Valor" _
    )
End Sub

Sub updateMemberPivotTable()
    Call updateMonthPivotTable( _
        ActiveSheet, _
        Defs.PIVOT_TABLE_MEMBRO_NOME, _
        obterRangeDespesasAnalise(ActiveSheet), _
        Defs.PIVOT_TABLE_MEMBRO_LINHA, _
        Defs.PIVOT_TABLE_MEMBRO_COLUNA, _
        "Membro", _
        "Valor" _
    )
End Sub

Function createExpenseReportPivotTable() As PivotTable
    Dim data As Range, dest As Range, sh As Worksheet, table As PivotTable
    Set sh = Sheets(Defs.SHEET_CONSOLIDATE)
    Set data = obterRangeTabela(sh, _
                                Defs.CONSOLIDATE_EXPENSE_START_LINE - 1, _
                                Defs.CONSOLIDATE_EXPENSE_START_COL, _
                                4)
    Set dest = sh.Cells(2, 11)
    Set table = createNewPivotTable(data, dest, Defs.PIVOT_TABLE_TEMP_EXPENSE)
        
    With table
        .PivotFields("Mês").Orientation = xlColumnField
        .PivotFields("Mês").ShowAllItems = True
        .PivotFields("Categoria").Orientation = xlRowField
        .PivotFields("Valor").Orientation = xlDataField
        .PivotFields("Categoria").LabelRange = "Categoria"
        .PivotFields("Categoria").PivotItems("(blank)").Visible = False
        .DataBodyRange.NumberFormat = "$ #,##0.00"
        .HasAutoFormat = False
        .RowGrand = False
        
    End With
    Call stylizeTable(table)
    
    Set createExpenseReportPivotTable = table
End Function

Private Sub updateMonthPivotTable(wksh As Worksheet, _
                                  tableName As String, _
                                  data As Range, _
                                  posRow As Integer, _
                                  posCol As Integer, _
                                  rowField As String, _
                                  dataField As String)
                                  
    Dim table As PivotTable, exist As Boolean
    exist = pivotExist(wksh, tableName)
    
    If (exist = True And Not data Is Nothing) Then
        Dim cache As PivotCache
        Set table = wksh.PivotTables(tableName)
        Set cache = createPivotCache(data)
        table.ChangePivotCache cache
    
    ElseIf (exist = True) Then
        wksh.PivotTables(tableName).TableRange2.Clear
        
    ElseIf (Not data Is Nothing) Then
        Dim dest As Range
        Set dest = wksh.Cells(posRow, posCol)
        Set table = createNewPivotTable(data, dest, tableName)
        
        With table
            .HasAutoFormat = False
            .PivotFields(rowField).Orientation = xlRowField
            .PivotFields(dataField).Orientation = xlDataField
            .PivotFields(rowField).LabelRange = rowField
            .DataBodyRange.NumberFormat = "$ #,##0.00"
            .ColumnGrand = False
        End With
        Call stylizeTable(table)
    End If
End Sub


Private Function createPivotCache(data As Range) As PivotCache
    Set createPivotCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=data)
End Function

Private Function createNewPivotTable(data As Range, destination As Range, name As String) As PivotTable
    Dim cache As PivotCache
    Set cache = createPivotCache(data)
    Set createNewPivotTable = cache.CreatePivotTable(TableDestination:=destination, tableName:=name)
End Function



Private Sub stylizeTable(table As PivotTable)
    
    With table.TableRange1
        .Borders.Color = vbWhite
        .Borders.Weight = xlMedium
        .Interior.Color = RGB(245, 226, 169)
        .HorizontalAlignment = xlCenter
        
        With .Rows(1)
            .Interior.Color = RGB(176, 74, 14)
            .Font.Color = RGB(255, 255, 255)
        End With
    End With
End Sub


Private Function pivotExist(sh As Worksheet, name As String) As Boolean
    Dim Pt As PivotTable
    pivotExist = False
    For Each Pt In sh.PivotTables
        If Pt.name = name Then
            pivotExist = True
            Exit For
        End If
    Next
End Function


