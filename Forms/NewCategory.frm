VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewCategory 
   Caption         =   "Nova Categoria"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4635
   OleObjectBlob   =   "NewCategory.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonAdd_Click()
    Dim tp As RecordType, cat As category
    If (OptionButtonIncome.value = True) Then
        tp = INCOME
    Else
        tp = EXPENSE
    End If
    
    On Error GoTo ShowError
    Call insertCategory(validateInputCategory(tp, TextBoxName.value))
    Unload NewCategory
    Exit Sub
    
ShowError:
    MsgBox Err.Description
End Sub
