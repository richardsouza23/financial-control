Attribute VB_Name = "Types"
Public Enum RecordType
    INCOME
    EXPENSE
End Enum

Public Type category
    name As String
    type As RecordType
End Type

Public Type RegistroValor
    diaDoMes As Integer
    membro As String
    Categoria As String
    descricao As String
    valor As Double
    type As RecordType
End Type


Public Type MonthRecord
    record As RegistroValor
    month As String
End Type
