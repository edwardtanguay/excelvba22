Option Explicit

Function qdat_GermanDate(TheDate As Date) As String
    qdat_GermanDate = qstr_PadTwoZeros(Day(TheDate)) & "." & qstr_PadTwoZeros(Month(TheDate)) & "." & Year(TheDate)
End Function

Function qdat_AmericanDate(TheDate As Date) As String
    qdat_AmericanDate = qstr_PadTwoZeros(Month(TheDate)) & "/" & qstr_PadTwoZeros(Day(TheDate)) & "/" & Year(TheDate)
End Function

Function qdat_ISO8601(TheDate As Date) As String
    qdat_ISO8601 = Year(TheDate) & "-" & qdat_PaddedMonth(TheDate) & "-" & qdat_PaddedDay(TheDate)
End Function

Function qdat_PaddedMonth(TheDate As Date) As String
    qdat_PaddedMonth = qstr_PadTwoZeros(Month(TheDate))
End Function

Function qdat_PaddedDay(TheDate As Date) As String
    qdat_PaddedDay = qstr_PadTwoZeros(Day(TheDate))
End Function
