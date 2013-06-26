' Функция GetUrlFromHyperlink - извлечение URL из гиперссылки
' Параметры:
' range - ячейка с гиперссылкой
'
' Код функции необходимо вставить в личную книгу макросов (HOWTO - https://github.com/mevoronin/ExcelMacroses/blob/master/PersonalMacrosBook.html)
'
' (c) Voronin Mihail, 2011 (http://chenado.net/1149.html)

Attribute VB_Name = "GetUrlFromHyperlink"
Public Function GetUrlFromHyperlink(ByVal range As range) As String
    If (range.Hyperlinks.Count > 0) Then
        GetUrlFromHyperlink = range.Hyperlinks(1).Address
    Else
        GetUrlFromHyperlink = ""
    End If
End Function
