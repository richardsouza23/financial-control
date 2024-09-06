Attribute VB_Name = "Calculos"
Sub atualizarReceitaTotal()
    Dim rangeReceitas, soma As Double
    Set rangeReceitas = obterRangeReceitas(ActiveSheet)
    soma = 0
    If (Not rangeReceitas Is Nothing) Then
        soma = WorksheetFunction.Sum(rangeReceitas.Columns(6))
    End If
    ActiveSheet.Cells(Defs.REC_TOTAL_LINHA, Defs.REC_TOTAL_COLUNA).value = soma
End Sub

Sub atualizarDespesaTotal()
    Dim rangeDespesas, soma As Double
    Set rangeDespesas = obterRangeDespesas(ActiveSheet)
    soma = 0
    If (Not rangeDespesas Is Nothing) Then
        soma = WorksheetFunction.Sum(rangeDespesas.Columns(6))
    End If
    ActiveSheet.Cells(Defs.DESP_TOTAL_LINHA, Defs.DESP_TOTAL_COLUNA).value = soma
End Sub


Function obterLinhaInicioDespesas(mSheet As Worksheet) As Integer
    Dim celRec As Range
    Set celRec = mSheet.Cells(Defs.INICIO_RECEITA_LINHA, Defs.INICIO_RECEITA_COLUNA)
    
    If (IsEmpty(mSheet.Cells(celRec.Row + 1, celRec.Column)) = True) Then
        obterLinhaInicioDespesas = celRec.End(xlDown).Row + 2
    Else
        obterLinhaInicioDespesas = celRec.End(xlDown).Row + 5
    End If
End Function

Function obterLinhaFinalTabela(planilha As Worksheet, linhaInicial As Integer, colunaInicial As Integer) As Integer
    Dim celIni As Range
    Set celIni = planilha.Cells(linhaInicial, colunaInicial)
    
    If (IsEmpty(celIni.value) = True) Then
        obterLinhaFinalTabela = linhaInicial - 1
    ElseIf (IsEmpty(planilha.Cells(linhaInicial + 1, colunaInicial).value) = True) Then
        obterLinhaFinalTabela = linhaInicial
    Else
        obterLinhaFinalTabela = celIni.End(xlDown).Row
    End If
End Function


Function obterRangeReceitas(planilha As Worksheet) As Range
    Set obterRangeReceitas = obterRangeTabela(planilha, _
                                              Defs.INICIO_RECEITA_LINHA, _
                                              Defs.INICIO_RECEITA_COLUNA, _
                                              6)
End Function

Function obterRangeReceitasAnalise(planilha As Worksheet) As Range
    Dim linhaInicio As Integer, linhaFim As Integer, rg As Range
    Set rg = obterRangeTabela(planilha, Defs.INICIO_RECEITA_LINHA - 1, Defs.INICIO_RECEITA_COLUNA + 1, 5)
    
    If (rg.Rows.Count = 1) Then
        Set obterRangeReceitasAnalise = Nothing
        Exit Function
    End If
    Set obterRangeReceitasAnalise = rg
End Function

Function obterRangeDespesas(planilha As Worksheet) As Range
    Set obterRangeDespesas = obterRangeTabela(planilha, _
                                              obterLinhaInicioDespesas(planilha), _
                                              Defs.INICIO_RECEITA_COLUNA, _
                                              6)
End Function

Function obterRangeDespesasAnalise(planilha As Worksheet) As Range
    Dim linhaInicio As Integer, linhaFim As Integer, rg As Range
    linhaInicio = obterLinhaInicioDespesas(planilha)
    Set rg = obterRangeTabela(planilha, linhaInicio - 1, Defs.INICIO_RECEITA_COLUNA + 1, 5)

    If (rg.Rows.Count = 1) Then
        Set obterRangeDespesasAnalise = Nothing
        Exit Function
    End If
    
    Set obterRangeDespesasAnalise = rg
End Function

Function obterRangeTabela(planilha As Worksheet, linhaInicial As Integer, colInicial As Integer, qtdColunas As Integer) As Range
    Dim linhaInicio As Integer
    linhaFim = obterLinhaFinalTabela(planilha, linhaInicial, colInicial)
    
    If (linhaInicial > linhaFim) Then
        Set obterRangeTabela = Nothing
        Exit Function
    End If
    
    Set obterRangeTabela = planilha.Range(planilha.Cells(linhaInicial, colInicial), planilha.Cells(linhaFim, colInicial + qtdColunas - 1))
End Function


