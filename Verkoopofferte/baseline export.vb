Sub ExportRangeToCSV()

    Dim ExportSheet As Worksheet
    Dim ExportRange As Range
    Dim BasePath As String
    Dim OfferNumber As String
    Dim FileName As String
    Dim FullPath As String
    Dim FileNum As Integer
    Dim Row As Long, Col As Long
    Dim Line As String

    ' === CONFIGURATION ===
    Set ExportSheet = ThisWorkbook.Sheets("Sheet1")
    Set ExportRange = ExportSheet.Range("A1:F100")
    BasePath = "W:\Almacon Offertes\Ridder Offertes\"
    OfferNumber = CleanFileName(ExportSheet.Range("B2").Value)
    FileName = CleanFileName(ExportSheet.Range("B3").Value)
    ' =====================

    If OfferNumber = "" Or FileName = "" Then
        MsgBox "Offer number or filename is missing.", vbExclamation
        Exit Sub
    End If

    ' Build full path
    FullPath = BasePath & OfferNumber & "\"

    ' Ensure folder exists
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
