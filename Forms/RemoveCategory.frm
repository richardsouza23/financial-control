VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveCategory 
   Caption         =   "Excluir Categoria"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4455
   OleObjectBlob   =   "RemoveCategory.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemoveCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonConfirm_Click()
    Dim value As Integer, rt As RecordType
    On Error GoTo ShowError
    value = converterParaInteiro(TextBoxLine.value)
    
    If (OptionButtonIncome.value = True) Then
        rt = INCOME
    Else
        rt = EXPENSE
    End If
    
    Call deleteCategory(value, rt)
    Unload RemoveCategory
    Exit Sub

ShowError:
    MsgBox Err.Description
End Sub
