Attribute VB_Name = "Utils"
Function converterParaInteiro(str As String) As Integer
    If (IsNumeric(Trim(str)) = False) Then
        Err.Raise vbObjectError + 513, "", "'" & str & "' não é um número!"
    End If
    converterParaInteiro = CInt(Trim(str))
End Function

Sub formatRangeAnyRecord(rg As Range)
    With rg
        .Interior.Color = RGB(245, 226, 169)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
        .Borders.Color = vbWhite
        .Borders.Weight = xlMedium
        .HorizontalAlignment = xlCenter
    End With
End Sub
