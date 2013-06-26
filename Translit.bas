' Функция Translit - транслитерация значения в ячейке
' Параметры:
' range - ячейка с текстом для перевода
'
' Код функции необходимо вставить в личную книгу макросов (HOWTO - https://github.com/mevoronin/ExcelMacroses/blob/master/PersonalMacrosBook.html)
'
' (c) Voronin Mihail, 2011 (http://chenado.net/1131.html)



Public Function Translit(ByVal txt As String) As String
iRussianLower$ = "абвгдеёжзийклмнопрстуфхцчшщъыьэюя"
iTranslit = Array("", _
"a", "b", "v", _
"g", "d", "e", _
"yo", "zh", "z", _
"i", "i", "k", _
"l", "m", "n", _
"o", "p", "r", _
"s", "t", "u", _
"f", "kh", "c", _
"ch", "sh", "zch", _
"''", "'y", "'", _
"eh", "yu", "ya")
Dim result$, char$, newChar$, charIndex%, nextChar$, lastChar$
For i% = 1 To Len(txt)
char = Mid(txt, i, 1)
charIndex = InStr(1, iRussianLower, char, vbTextCompare)
If (charIndex >= 1) Then
newChar = iTranslit(charIndex)
Else
newChar = char
End If
' если текущий символ прописной
If (Asc(char) >= 132 And Asc(char) <= 223) Then
' если это первый символ
If i = 1 Then
' если последующий прописной
nextChar = Mid(txt, i + 1, 1)
If (Asc(nextChar) >= 132 And Asc(nextChar) <= 223) Then
result = result & StrConv(newChar, vbUpperCase)
' если последующий строчный
ElseIf (Asc(nextChar) >= 224 And Asc(nextChar) <= 255) Then
result = result & StrConv(newChar, vbProperCase)
End If
' если это не первый и не последний символ
ElseIf i > 1 And i <> Len(txt) Then
nextChar = Mid(txt, i + 1, 1)
lastChar = Mid(txt, i - 1, 1)
' если околостоящие прописные
If ((Asc(lastChar) >= 132 And Asc(lastChar) <= 223)) _
Or (Asc(nextChar) >= 132 And Asc(nextChar) <= 223) Then
result = result & StrConv(newChar, vbUpperCase)
' иначе
Else
result = result & StrConv(newChar, vbProperCase)
End If
Else
lastChar = Mid(txt, i - 1, 1)
' если предыдущий символ прописной
If (Asc(lastChar) >= 132 And Asc(lastChar) <= 223) Then
result = result & StrConv(newChar, vbProperCase)
' иначе
Else
result = result & StrConv(newChar, vbUpperCase)
End If
End If
' если текущий символ строчный
ElseIf (Asc(char) >= 224 And Asc(char) <= 255) Then
result = result & iTranslit(charIndex)
' если текущий символ не буква русского алфавита
Else
result = result & char
End If
Next i
Translit$ = result
End Function