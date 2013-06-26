Attribute VB_Name = "GAtags"
' Функция AddGAtags - добавление к ссылке меток Google Analytics
' Параметры:
' source_url - исходная ссылка
' utm_source - источник
' utm_campaign - название кампании
' utm_medium - средство кампании
' utm_term - ключевое слово
' utm_content - содержание кампании
'
' (c) Voronin Mihail, 2011 (http://chenado.net/1163.html)

Option Explicit

Type UrlParamsParts
    paramKey As String
    paramValue As String
End Type

Public Function AddGAtags(ByVal source_url As Variant, Optional ByVal utm_source As Variant, Optional ByVal utm_campaign As Variant, Optional ByVal utm_medium As Variant, Optional ByVal utm_term As Variant = "", Optional ByVal utm_content As Variant = "") As String
    Dim url As String, page As String, query_string As String, i As Integer
    Dim params() As UrlParamsParts
    Dim anchor As String
    Dim result As String
    anchor = ""
    url = LCase(source_url)
    If (InStr(1, url, "?") > 0) Then
        page = Mid(url, 1, InStr(1, url, "?") - 1)
        If InStr(1, page, "#") > 0 Then page = Mid(page, 1, InStr(1, url, "#") - 1)
        query_string = Mid(url, InStr(1, url, "?") + 1)
    Else
        page = url
    End If
    Dim sharpPosition As Integer
    sharpPosition = InStr(1, query_string, "#")
    If (sharpPosition > 0) Then
        anchor = Mid(query_string, sharpPosition)
        query_string = Mid(query_string, 1, InStr(1, query_string, "#") - 1)
    Else
        sharpPosition = InStr(1, url, "#")
        If (sharpPosition > 0) Then
            anchor = Mid(url, sharpPosition)
            Dim questionPosition As Integer
            questionPosition = InStr(1, anchor, "?")
            If (questionPosition > 0) Then
                query_string = Mid(anchor, questionPosition + 1)
                anchor = Mid(anchor, 1, questionPosition - 1)
            End If
        End If
    End If
    
    SplitQueryString query_string, params
    
    result = page
    If IsMissing(utm_source) = False Then
        If utm_source <> "" Then AddParam params, "utm_source", CStr(utm_source)
    End If
    If IsMissing(utm_campaign) = False Then
        If utm_campaign <> "" Then AddParam params, "utm_campaign", CStr(utm_campaign)
    End If
    If IsMissing(utm_medium) = False Then
        If utm_medium <> "" Then AddParam params, "utm_medium", CStr(utm_medium)
    End If
    If (UBound(params) > 0) Then
        result = result & "?"
        For i = 1 To UBound(params)
            result = result & params(i).paramKey & "=" & params(i).paramValue & "&"
        Next i
        result = Mid(result, 1, Len(result) - 1)
    End If
    If (anchor <> "") Then result = result & anchor
    AddGAtags = result
End Function

Private Sub AddParam(ByRef params() As UrlParamsParts, paramKey As String, paramValue As String)
    Dim i As Integer
    For i = 1 To UBound(params)
        If (params(i).paramKey = paramKey) Then
            params(i).paramValue = paramValue
            Exit Sub
        End If
    Next i
    ReDim Preserve params(UBound(params) + 1)
    params(UBound(params)).paramKey = paramKey
    params(UBound(params)).paramValue = paramValue
    
End Sub

Private Sub SplitQueryString(ByVal query_string As String, ByRef params() As UrlParamsParts)
    Dim i As Integer, paramKey As String, paramValue As String
    Dim lastStart As Integer
    ReDim params(1)
    ReDim Preserve params(UBound(params) - 1)
    If (query_string = "") Then Exit Sub
    lastStart = 1
    For i = 1 To Len(query_string)
        If (Mid(query_string, i, 1) = "=") Then
            paramKey = Mid(query_string, lastStart, i - lastStart)
            lastStart = i + 1
        End If
        If (Mid(query_string, i, 1) = "&" Or Mid(query_string, i, 1) = "?" Or i = Len(query_string)) Then
            If (i = Len(query_string)) Then
                paramValue = Mid(query_string, lastStart)
            Else
                paramValue = Mid(query_string, lastStart, i - lastStart)
            End If
            lastStart = i + 1
            AddParam params, paramKey, paramValue
        End If
    Next i
    
End Sub


