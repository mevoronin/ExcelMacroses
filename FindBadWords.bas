Attribute VB_Name = "FindBadWords"
Public Function FindBadWords(title As String, text As String) As Integer
Dim words(1 To 10) As String
words(1) = "скидки"
words(2) = "лучший"
words(3) = "бесплатно"
words(4) = "акция"
words(5) = "распродажа"
words(6) = "скидка"
words(7) = "бесплатный"
words(8) = "лучше"
words(9) = "акции"
words(10) = "распродажи"
Dim i As Integer
For i = 1 To UBound(words)
    If (InStr(1, LCase(title), words(i)) >= 0) Or (InStr(1, LCase(text), words(i)) >= 0) Then
        FindBadWords = "ИСТИНА"
        Exit Function
    End If
    
Next i
FindBadWords = "ЛОЖЬ"
End Function
