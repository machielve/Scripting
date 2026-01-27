Sub ExportRangeToCSV()

    Dim ExportSheet As Worksheet
    Dim ExportRange As Range
    Dim BasePath As String
    Dim OrderNumber As Long
    Dim FileName As String
    Dim RangeFolder As String
    Dim FullPath As String
    Dim FileNum As Integer
    Dim Row As Long, Col As Long
    Dim Line As String

    ' === CONFIGURATION ===
    Set ExportSheet = ThisWorkbook.Sheets("Sheet1")
    Set ExportRange = ExportSheet.Range("A1:F100")
    BasePath = "W:\Almacon Offertes\Ridder Offertes\"
    OrderNumber = ExportSheet.Range("B2").Value      ' 24667
    FileName = CleanFileName(ExportSheet.Range("B3").Value) ' 24667-1
    ' =====================

    If FileName = "" Or OrderNumber = 0 Then
        MsgBox "Order number or filename missing.", vbExclamation
        Exit Sub
    End If

    ' Build range folder (24600-24699)
    RangeFolder = (OrderNumber \ 100) * 100 & "-" & ((OrderNumber \ 100) * 100 + 99)

    ' Build full path
    FullPath = BasePath & RangeFolder & "\" & OrderNumber & "\"

    ' Ensure folders exist
    CreateFolderIfMissing BasePath & RangeFolder
    CreateFolderIfMissing FullPath

    FileNum = FreeFile
    Open FullPath & FileName & ".csv" For Output As #FileNum

    For Row = 1 To ExportRange.Rows.Count
        Line = ""
        For Col = 1 To ExportRange.Columns.Count
            Line = Line & """" & ExportRange.Cells(Row, Col).Text & """"
            If Col < ExportRange.Columns.Count Then Line = Line & ";"
        Next Col
        Print #FileNum, Line
    Next Row

    Close #FileNum

    MsgBox "CSV exported to:" & vbCrLf & FullPath & FileName & ".csv", vbInformation

End Sub
