Attribute VB_Name = "GetUrlFromHyperlink"
Public Function GetUrlFromHyperlink(ByVal range As range) As String
    If (range.Hyperlinks.Count > 0) Then
        GetUrlFromHyperlink = range.Hyperlinks(1).Address
    Else
        GetUrlFromHyperlink = ""
    End If
End Function
