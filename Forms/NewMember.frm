VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewMember 
   Caption         =   "Novo Membro"
   ClientHeight    =   2550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4485
   OleObjectBlob   =   "NewMember.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonAdd_Click()
    Dim tp As RecordType, cat As category
    If (Trim(TextBoxName.value) = "") Then
        MsgBox "Nome inválido!"
    End If
    
    On Error GoTo ShowError
    Call insertMember(Trim(TextBoxName.value))
    Unload NewMember
    Exit Sub
    
ShowError:
    MsgBox Err.Description
End Sub
