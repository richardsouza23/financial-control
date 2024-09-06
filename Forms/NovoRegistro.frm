VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NovoRegistro 
   Caption         =   "Novo Registro"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
   OleObjectBlob   =   "NovoRegistro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NovoRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub botaoSalvarForm_Click()
    Dim regValor As RegistroValor, inicioCelular As Range
    
    On Error GoTo MostrarErro
    regValor = validarInputRegistro(TextBoxDiaMes.value, _
                                    ComboBoxMembro.value, _
                                    ComboBoxCategoria.value, _
                                    TextBoxDescricao.value, _
                                    TextBoxValor.value)
    
    If (OptionButtonReceita.value = True) Then
        Call inserirReceita(regValor)
    Else
        Call inserirDespesa(regValor)
    End If
    
    Unload NovoRegistro
    Exit Sub
    
MostrarErro:
    MsgBox Err.Description
End Sub


Private Sub OptionButtonDespesa_Click()
    Call atualizarCategorias
End Sub

Private Sub OptionButtonReceita_Click()
    Call atualizarCategorias
End Sub

Private Sub UserForm_Initialize()
    Call atualizarCategorias
    Call preencherMembros(ComboBoxMembro, Defs.INICIO_MEMBROS_COLUNA + 1, Defs.INICIO_MEMBROS_LINHA)
End Sub


Private Sub atualizarCategorias()
    If (OptionButtonReceita.value = True) Then
        Call preencherCategorias(ComboBoxCategoria, Defs.INICIO_CATEGORIAS_RECEITA_COLUNA + 1, Defs.INICIO_CATEGORIAS_RECEITA_LINHA)
    Else
        Call preencherCategorias(ComboBoxCategoria, Defs.INICIO_CATEGORIAS_DESPESA_COLUNA + 1, Defs.INICIO_CATEGORIAS_DESPESA_LINHA)
    End If
    ComboBoxCategoria.value = ""
End Sub
