Attribute VB_Name = "FindBadWords"
Public Function FindBadWords(title As String, text As String) As Integer
Dim words(1 To 10) As String
words(1) = "������"
words(2) = "������"
words(3) = "���������"
words(4) = "�����"
words(5) = "����������"
words(6) = "������"
words(7) = "����������"
words(8) = "�����"
words(9) = "�����"
words(10) = "����������"
Dim i As Integer
For i = 1 To UBound(words)
    If (InStr(1, LCase(title), words(i)) >= 0) Or (InStr(1, LCase(text), words(i)) >= 0) Then
        FindBadWords = "������"
        Exit Function
    End If
    
Next i
FindBadWords = "����"
End Function
