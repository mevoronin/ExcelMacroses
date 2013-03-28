' Макрос для удаления повторяющихся строк.
' Удаление производится в первом стобце. Данные должны быть отсортированы по возарстанию

Public Sub DeleteNotUnicalRows()
Dim i%
i = 1
Dim deleted As Integer
deleted = 0
If MsgBox("Вы действительно хотите удалить все повторяющиеся записи?", vbYesNo, "Удаление повторов") = vbNo Then Exit Sub

Do While Not IsEmpty(ActiveSheet.Cells(i, 1))
    Do While (UCase(Trim(ActiveSheet.Cells(i, 1))) = UCase(Trim(ActiveSheet.Cells(i + 1, 1))))
        ActiveSheet.range(CStr("A" & i)).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 5296274
        End With
        ActiveSheet.rows(CStr(i + 1) & ":" & CStr(i + 1)).Select
        Selection.Delete Shift:=xlUp
        deleted = deleted + 1
    Loop
    
    i = i + 1
Loop
MsgBox "Удалено " & deleted & " строк"

End Sub

