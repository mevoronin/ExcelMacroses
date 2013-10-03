Public Sub JoinColumnsWithWordsIntoOneColumn()
' Combines the columns with the words in one column.
' Name of subjects will be listed in the next column

Dim currentRowIndex%, newRowIndex%
Dim currentColumnIndex%
Dim curSheet As Worksheet, newSheet As Worksheet
Dim thematicName$
Set curSheet = ActiveSheet
Set newSheet = Sheets.Add()
newSheet.Name = "Joined words"
currentRowIndex = 0
currentColumnIndex = 0
newRowIndex = 0

Do While (Not IsEmpty(curSheet.Cells(1, currentColumnIndex + 1)))
    currentColumnIndex = currentColumnIndex + 1
    currentRowIndex = 1
    thematicName = curSheet.Cells(currentRowIndex, currentColumnIndex)
    Do While (Not IsEmpty(curSheet.Cells(currentRowIndex + 1, currentColumnIndex)))
        currentRowIndex = currentRowIndex + 1
        newRowIndex = newRowIndex + 1
        newSheet.Cells(newRowIndex, 1) = thematicName
        newSheet.Cells(newRowIndex, 2) = curSheet.Cells(currentRowIndex, currentColumnIndex)
    Loop
Loop

End Sub

