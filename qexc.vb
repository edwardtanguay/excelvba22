Public Sub qexc_HighlightCell(cell As Range)
    With cell.Interior
        .Pattern = xlSolid
        .PatternColor = 16777215
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    cell.Font.Bold = True
End Sub
 
Public Sub qexc_UnhighlightCell(cell As Range)
    With cell.Interior
        .Pattern = xlSolid
        .PatternColor = 16777215
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    cell.Font.Bold = False
End Sub

Public Sub qexc_GotoWorksheet(theWorksheet As String)
    Sheets(theWorksheet).Activate
    Sheets(theWorksheet).Range("A1").Select
End Sub

