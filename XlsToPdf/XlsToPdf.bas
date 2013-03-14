Attribute VB_Name = "XlsToPdf"
Sub XlsToPdf()
Dim path$, file$, test$
Dim reportIndex As Integer
Dim book As Workbook
path = ActiveWorkbook.path & "\"
file = Dir(path & "*.xls")
reportIndex = 0
Do While file <> ""
If (Right(file, 5) <> ".xlsm") Then
reportIndex = reportIndex + 1
Cells(reportIndex, 1) = file
file = path & file
Dim rowIndex As Integer
Set book = Workbooks.Open(file, False, True)
    If (InStr(1, book.ActiveSheet.Cells(4, 2), "Àêò ¹") > 0) Then
        rowIndex = 8
        Do While Not IsEmpty(book.ActiveSheet.Cells(rowIndex + 1, 2))
            rowIndex = rowIndex + 1
            Rows(CStr(rowIndex & ":" & rowIndex)).RowHeight = Rows(CStr(rowIndex & ":" & rowIndex)).RowHeight + 20
        Loop
    ElseIf (InStr(1, book.ActiveSheet.Cells(5, 2), "Ñ÷åò-ôàêòóðà") > 0) Then
        rowIndex = 18
        Do While Not IsEmpty(book.ActiveSheet.Cells(rowIndex + 1, 2))
            rowIndex = rowIndex + 1
            Rows(CStr(rowIndex & ":" & rowIndex)).RowHeight = Rows(CStr(rowIndex & ":" & rowIndex)).RowHeight + 20
        Loop
    ElseIf (InStr(1, book.ActiveSheet.Cells(12, 2), "Ñ×ÅÒ ¹") > 0) Then
        rowIndex = 16
        Do While Not IsEmpty(book.ActiveSheet.Cells(rowIndex + 1, 2))
            rowIndex = rowIndex + 1
            Rows(CStr(rowIndex & ":" & rowIndex)).RowHeight = Rows(CStr(rowIndex & ":" & rowIndex)).RowHeight + 20
        Loop
    End If
    With book.ActiveSheet.PageSetup
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        test = book.ActiveSheet.Cells(4, 2)
    If (InStr(1, book.ActiveSheet.Cells(5, 2), "Ñ÷åò-ôàêòóðà") > 0) Then
        .Orientation = xlLandscape
        Columns("C:C").ColumnWidth = 7.83
        Rows("2:2").RowHeight = 42
    Else
        .Orientation = xlPortrait
    End If
    If (InStr(1, book.ActiveSheet.Cells(12, 2), "Ñ×ÅÒ ¹") > 0) Then
        Columns("E:E").ColumnWidth = 10
    End If
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
book.ExportAsFixedFormat Type:=xlTypePDF, Filename:=file & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
book.Close (False)
End If
file = Dir
Loop
End Sub

