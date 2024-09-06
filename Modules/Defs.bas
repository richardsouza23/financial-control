Attribute VB_Name = "Defs"
Public Const PLANILHA_CATEGORIAS As String = "Categorias"
Public Const PLANILHA_MEMBROS = "Membros"
Public Const SHEET_CONSOLIDATE As String = "consolidate"
Public Const SHEET_REPORT As String = "Resumo"

Public Const INICIO_CATEGORIAS_RECEITA_COLUNA As Integer = 2
Public Const INICIO_CATEGORIAS_RECEITA_LINHA As Integer = 8

Public Const INICIO_CATEGORIAS_DESPESA_COLUNA As Integer = 5
Public Const INICIO_CATEGORIAS_DESPESA_LINHA As Integer = 8

Public Const INICIO_MEMBROS_COLUNA As Integer = 2
Public Const INICIO_MEMBROS_LINHA As Integer = 7

Public Const INICIO_RECEITA_COLUNA As Integer = 2
Public Const INICIO_RECEITA_LINHA As Integer = 7
Public Const QTD_COLUNAS_RECEITA As Integer = 6

Public Const QTD_COLUNAS_DESPESA As Integer = 6

Public Const REC_TOTAL_COLUNA As Integer = 10, REC_TOTAL_LINHA As Integer = 2
Public Const DESP_TOTAL_COLUNA As Integer = 10, DESP_TOTAL_LINHA As Integer = 4

Public Const PIVOT_TABLE_CATEG_LINHA As Integer = 11, PIVOT_TABLE_CATEG_COLUNA As Integer = 9
Public Const PIVOT_TABLE_MEMBRO_LINHA As Integer = 11, PIVOT_TABLE_MEMBRO_COLUNA As Integer = 12

Public Const PIVOT_TABLE_CATEG_NOME As String = "Despesas por Categoria"
Public Const PIVOT_TABLE_MEMBRO_NOME As String = "Despesas por Membro"
Public Const PIVOT_TABLE_TEMP_EXPENSE = "Expenses Report Temp"

Public Const CELULA_NOME_MES As String = "B2"

Public Const CONSOLIDATE_EXPENSE_START_LINE As Integer = 3
Public Const CONSOLIDATE_EXPENSE_START_COL As Integer = 6

Public Const REPORT_EXPENSE_START_LINE As Integer = 12
Public Const REPORT_EXPENSE_START_COL As Integer = 2


Public Const SHEET_ROWCOUNT As Long = 1048576
Public Const REPORT_FULL_RANGE As String = "B6:Q1000"


Function getMonthSheets() As Variant()
    getMonthSheets = Array("JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ")
End Function
