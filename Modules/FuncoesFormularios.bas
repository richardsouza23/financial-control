Attribute VB_Name = "FuncoesFormularios"
Sub abrirFormularioNovoRegistro()
    NovoRegistro.Show
End Sub

Sub abrirFormularioExcluirRegistro()
    ExcluirRegistro.Show
End Sub

Sub abrirFormularioNovaCategoria()
    NewCategory.Show
End Sub

Sub abrirFormularioExcluirCategoria()
    RemoveCategory.Show
End Sub

Sub abrirFormularioNovoMembro()
    NewMember.Show
End Sub

Sub abrirFormularioExcluirMembro()
    RemoveMember.Show
End Sub

Sub preencherCategorias(comboBox As comboBox, colCelInicial As Integer, linCelInicial As Integer)
    Call preencherComboBox(comboBox, Defs.PLANILHA_CATEGORIAS, colCelInicial, linCelInicial)
End Sub

Sub preencherMembros(comboBox As comboBox, colCelInicial As Integer, linCelInicial As Integer)
    Call preencherComboBox(comboBox, Defs.PLANILHA_MEMBROS, colCelInicial, linCelInicial)
End Sub


Private Sub preencherComboBox(comboBox As comboBox, nomePlanilha As String, colInicial As Integer, linInicial As Integer)
    Dim finalLine As Integer, rg As Range
    Set rg = obterRangeTabela(Sheets(nomePlanilha), linInicial, colInicial, 1)
    If (Not rg Is Nothing) Then
        comboBox.RowSource = nomePlanilha & "!" & rg.Address(rowAbsolute:=False, columnAbsolute:=False)
    End If
End Sub


Sub insertCategory(cat As category)
    Dim startLine As Integer, finalLine As Integer, startCol As Integer, inserted As Boolean
    Dim rgTable As Range, mSheet As Worksheet
    
    Set mSheet = Sheets(Defs.PLANILHA_CATEGORIAS)
    inserted = False
    
    If (cat.type = INCOME) Then
        startLine = Defs.INICIO_CATEGORIAS_RECEITA_LINHA
        startCol = Defs.INICIO_CATEGORIAS_RECEITA_COLUNA
    Else
        startLine = Defs.INICIO_CATEGORIAS_DESPESA_LINHA
        startCol = Defs.INICIO_CATEGORIAS_DESPESA_COLUNA
    End If
    
    Set rgTable = obterRangeTabela(mSheet, startLine, startCol, 2)
    If (rgTable Is Nothing) Then
        finalLine = startLine
    Else
        finalLine = startLine + rgTable.Rows.Count
    End If
        
    For i = startLine To finalLine
        currentCat = mSheet.Cells(i, startCol + 1).value
            
        If (inserted = False And StrComp(currentCat, cat.name) = 0) Then
            Err.Raise vbObjectError + 513, "", "Categoria já existe!"
            Exit Sub
        End If
            
        If (inserted = False And (currentCat = "" Or StrComp(currentCat, cat.name) > 0)) Then
            Dim rgLine As Range
            mSheet.Range(Cells(i, startCol), Cells(i, startCol + 1)).Insert
            mSheet.Cells(i, startCol + 1).value = cat.name
            Call formatRangeAnyRecord(mSheet.Range(Cells(i, startCol), Cells(i, startCol + 1)))
            inserted = True
        End If
            
        If (inserted = True) Then
            mSheet.Cells(i, startCol).value = i - startLine + 1
        End If
    Next
End Sub

Sub insertMember(name As String)
    Dim startLine As Integer, finalLine As Integer, startCol As Integer, inserted As Boolean
    Dim rgTable As Range, mSheet As Worksheet
    
    Set mSheet = Sheets(Defs.PLANILHA_MEMBROS)
    inserted = False
    startLine = Defs.INICIO_MEMBROS_LINHA
    startCol = Defs.INICIO_MEMBROS_COLUNA
    Set rgTable = obterRangeTabela(mSheet, startLine, startCol, 2)
    
    If (rgTable Is Nothing) Then
        finalLine = startLine
    Else
        finalLine = startLine + rgTable.Rows.Count
    End If
        
    For i = startLine To finalLine
        current = mSheet.Cells(i, startCol + 1).value
            
        If (inserted = False And StrComp(current, name) = 0) Then
            Err.Raise vbObjectError + 513, "", "Membro já existe!"
            Exit Sub
        End If
            
        If (inserted = False And (current = "" Or StrComp(current, name) > 0)) Then
            Dim rgLine As Range
            mSheet.Range(Cells(i, startCol), Cells(i, startCol + 1)).Insert
            mSheet.Cells(i, startCol + 1).value = name
            Call formatRangeAnyRecord(mSheet.Range(Cells(i, startCol), Cells(i, startCol + 1)))
            inserted = True
        End If
            
        If (inserted = True) Then
            mSheet.Cells(i, startCol).value = i - startLine + 1
        End If
    Next
End Sub

Sub inserirReceita(regValor As RegistroValor)
    Set inicioCelula = ActiveSheet.Cells(Defs.INICIO_RECEITA_LINHA, Defs.INICIO_RECEITA_COLUNA)
    Call inserirRegistro(regValor, inicioCelula.Column, inicioCelula.Row)
    Call atualizarReceitaTotal
End Sub

Sub inserirDespesa(regValor As RegistroValor)
    Set celRec = ActiveSheet.Cells(Defs.INICIO_RECEITA_LINHA, Defs.INICIO_RECEITA_COLUNA)
    Call inserirRegistro(regValor, celRec.Column, obterLinhaInicioDespesas(ActiveSheet))
    Call atualizarDespesaTotal
    Call updateCategoryPivotTable
    Call updateMemberPivotTable
    Call generateReport
End Sub

Sub inserirRegistro(regValor As RegistroValor, colInicial As Integer, linhaInicial As Integer)
    Dim rangeTab As Range, rangeLinha As Range, linhaFinal As Integer
    inseriu = False
    
    If (IsEmpty(ActiveSheet.Cells(linhaInicial, colInicial)) = True) Then
        linhaFinal = linhaInicial
    ElseIf (IsEmpty(ActiveSheet.Cells(linhaInicial + 1, colInicial)) = True) Then
        linhaFinal = linhaInicial + 1
    Else
        linhaFinal = ActiveSheet.Cells(linhaInicial, colInicial).End(xlDown).Row + 1
    End If
 
    For i = linhaInicial To linhaFinal
        diaMesCorrente = ActiveSheet.Cells(i, colInicial + 1).value
        
        If (inseriu = False And (diaMesCorrente = "" Or diaMesCorrente > regValor.diaDoMes)) Then
            ActiveSheet.Range(Cells(i, colInicial), Cells(i, colInicial + 5)).Insert
            Set rangeLinha = ActiveSheet.Range(Cells(i, colInicial), Cells(i, colInicial + 5))
            
            With rangeLinha
                .Cells(1, 2).value = regValor.diaDoMes
                .Cells(1, 3).value = regValor.membro
                .Cells(1, 4).value = regValor.Categoria
                .Cells(1, 5).value = regValor.descricao
                .Cells(1, 6).value = regValor.valor
                .Cells(1, 6).NumberFormat = "$ #,##0.00"
            End With
            Call formatRangeAnyRecord(rangeLinha)
            inseriu = True
        End If
        
        ActiveSheet.Cells(i, colInicial).value = i - linhaInicial + 1
    Next
End Sub

Sub deleteMember(line As Integer)
    Dim startLine As Integer, startCol As Integer, finalLine As Integer
    startLine = Defs.INICIO_MEMBROS_LINHA
    startCol = Defs.INICIO_MEMBROS_COLUNA
    finalLine = obterLinhaFinalTabela(Sheets(Defs.PLANILHA_MEMBROS), startLine, startCol)
    
    If (line < 1 Or line > finalLine - startLine + 1) Then
        Err.Raise vbObjectError + 513, "", "Linha inválida!"
        Exit Sub
    End If
    
    Call apagarRegistro(startLine + line - 1, startCol, finalLine, 2)
End Sub

Sub deleteCategory(line As Integer, rt As RecordType)
    Dim startLine As Integer, startCol As Integer, finalLine As Integer
    If (rt = INCOME) Then
        startLine = Defs.INICIO_CATEGORIAS_RECEITA_LINHA
        startCol = Defs.INICIO_CATEGORIAS_RECEITA_COLUNA
    Else
        startLine = Defs.INICIO_CATEGORIAS_DESPESA_LINHA
        startCol = Defs.INICIO_CATEGORIAS_DESPESA_COLUNA
    End If
    
    finalLine = obterLinhaFinalTabela(Sheets(Defs.PLANILHA_CATEGORIAS), startLine, startCol)
    
    If (line < 1 Or line > finalLine - startLine + 1) Then
        Err.Raise vbObjectError + 513, "", "Linha inválida!"
        Exit Sub
    End If
    
    Call apagarRegistro(startLine + line - 1, startCol, finalLine, 2)
End Sub

Sub apagarReceita(linhaApagar As Integer)
    Dim linhaFinal As Integer
    linhaFinal = obterLinhaFinalTabela(ActiveSheet, Defs.INICIO_RECEITA_LINHA, Defs.INICIO_RECEITA_COLUNA)
    
    If (linhaApagar < 1 Or linhaApagar > linhaFinal - Defs.INICIO_RECEITA_LINHA + 1) Then
        Err.Raise vbObjectError + 513, "", "Linha inválida!"
        Exit Sub
    End If
    
    Call apagarRegistro(Defs.INICIO_RECEITA_LINHA + linhaApagar - 1, _
                        Defs.INICIO_RECEITA_COLUNA, _
                        linhaFinal, _
                        Defs.QTD_COLUNAS_RECEITA)
    Call atualizarReceitaTotal
End Sub


Sub apagarDespesa(linhaApagar As Integer)
    Dim linhaIniDespesas As Integer, linhaFinal As Integer
    linhaIniDespesas = obterLinhaInicioDespesas(ActiveSheet)
    linhaFinal = obterLinhaFinalTabela(ActiveSheet, linhaIniDespesas, Defs.INICIO_RECEITA_COLUNA)
    
    If (linhaApagar < 1 Or linhaApagar > linhaFinal - linhaIniDespesas + 1) Then
        MsgBox "Linha inválida!"
        Exit Sub
    End If
    
    Call apagarRegistro(linhaIniDespesas + linhaApagar - 1, _
                        Defs.INICIO_RECEITA_COLUNA, _
                        linhaFinal, _
                        Defs.QTD_COLUNAS_DESPESA)
    Call atualizarDespesaTotal
    Call updateCategoryPivotTable
    Call updateMemberPivotTable
End Sub


Sub apagarRegistro(linha As Integer, coluna As Integer, linhaFinal As Integer, qtdColunas As Integer)
    linhaTabela = ActiveSheet.Cells(linha, coluna).value
    
    If (IsNumeric(Trim(linhaTabela)) = True) Then
        intNumLinha = CInt(Trim(linhaTabela))
        ActiveSheet.Range(Cells(linha, coluna), Cells(linha, coluna + qtdColunas - 1)).Delete shift:=xlUp
        
        For i = linha To linhaFinal - 1
            ActiveSheet.Cells(i, coluna).value = linhaTabela
            linhaTabela = linhaTabela + 1
        Next
    End If
End Sub

Function validarInputRegistro(dia As String, membro As String, Categoria As String, descricao As String, valor As String) As RegistroValor
    Dim RegistroValor As RegistroValor
    
    If (IsNumeric(Trim(dia)) = False) Then
        Err.Raise vbObjectError + 513, "", "Dia do mês inválido!"
    End If
    diaMes = CInt(Trim(dia))
    If (diaMes < 1 Or diaMes > 31) Then
        Err.Raise vbObjectError + 513, "", "Dia do mês inválido!"
    End If
    RegistroValor.diaDoMes = diaMes
    
    membro = Trim(membro)
    If (membro = "") Then
        Err.Raise vbObjectError + 513, "", "Membro inválido!"
    End If
    RegistroValor.membro = membro
    
    Categoria = Trim(Categoria)
    If (Categoria = "") Then
        Err.Raise vbObjectError + 513, "", "Categoria inválida!"
    End If
    RegistroValor.Categoria = Categoria

    RegistroValor.descricao = Trim(descricao)

    If (IsNumeric(Trim(valor)) = False) Then
        Err.Raise vbObjectError + 513, "", "Valor inválido!"
    End If
    valor = CDbl(Trim(valor))
    If (valor < 0.01) Then
        Err.Raise vbObjectError + 513, "", "Valor inválido!"
    End If
    RegistroValor.valor = valor
    
    validarInputRegistro = RegistroValor
    
End Function

Function validateInputCategory(rt As RecordType, name As String) As category
    If (Trim(name) = "") Then
        Err.Raise vbObjectError + 513, "", "Categoria inválida!"
    End If
    
    Dim cat As category
    cat.type = rt
    cat.name = Trim(name)
    validateInputCategory = cat
End Function


