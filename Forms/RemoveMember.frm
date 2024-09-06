VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveMember 
   Caption         =   "Excluir Membro"
   ClientHeight    =   2025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4455
   OleObjectBlob   =   "RemoveMember.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemoveMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonConfirm_Click()
    Dim value As Integer
    On Error GoTo ShowError
    value = converterParaInteiro(TextBoxLine.value)
    
    Call deleteMember(value)
    Unload RemoveMember
    Exit Sub

ShowError:
    MsgBox Err.Description
End Sub
