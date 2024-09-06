VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExcluirRegistro 
   Caption         =   "Excluir Registro"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4710
   OleObjectBlob   =   "ExcluirRegistro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExcluirRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonConfirmar_Click()
    Dim valor As Integer
    On Error GoTo MostrarErro
    valor = converterParaInteiro(TextBoxNumLinha.value)
    
    If (OptionButtonReceita.value = True) Then
        Call apagarReceita(valor)
    Else
        Call apagarDespesa(valor)
    End If
    
    Unload ExcluirRegistro
    Exit Sub

MostrarErro:
    MsgBox Err.Description
End Sub

