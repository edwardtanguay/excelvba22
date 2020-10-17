Function qfil_CreateTestFile(rngMain As Range) As String
    Dim lngFileNumber As Long
    'Dim rngMain As Range
    
    'rngMain = Range("D16:D18")
    lngFileNumber = FreeFile()
    
    Open ThisWorkbook.Path & "\test.txt" For Output As #lngFileNumber
    Print #lngFileNumber, "This is the data:"
    
    For Each rngMain In Selection.Cells
        Print #lngFileNumber, rngMain.Value & ""
    Next
    
    Close #lngFileNumber
End Function

Function qfil_CreateFile(pathAndFileName As String, content As String)
    Dim lngFileNumber As Long
    lngFileNumber = FreeFile()
    
    Open pathAndFileName For Output As #lngFileNumber
    Print #lngFileNumber, content
    
    Close #lngFileNumber
End Function
