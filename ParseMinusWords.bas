' Макрос ParseMinusWords - разбиение минус-слов по строчкам
' Входные данные - строка - в ячейке [1,1]
' (c) Voronin Mikhail, 2011
Attribute VB_Name = "ParseMinusWords"
Public Sub ParseMinusWords()
Dim str As String
str = ActiveSheet.Cells(1, 1)
Dim buffer As String
Dim resultRowIndex As Integer
Dim i As Integer
Dim lookChar As Boolean
resultRowIndex = 1
For i = 1 To Len(str)
    If (Mid(str, i, 1) = "-") Then lookChar = True
    If (Mid(str, i, 1) <> "-" And lookChar = True) Then
        buffer = buffer & Mid(str, i, 1)
    End If
    If ((buffer <> "" And Mid(str, i, 1) = "-") Or i = Len(str)) Then
        resultRowIndex = resultRowIndex + 1
        ActiveSheet.Cells(resultRowIndex, 1) = Trim(buffer)
        buffer = ""
    End If
Next i
End Sub
