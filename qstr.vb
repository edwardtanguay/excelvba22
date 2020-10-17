Function qstr_ChopRight(strContent As String, strPieceToChop As String) As String
    
    'declarations
    Dim r As String
    Dim intLengthOfPieceToChop As Integer
    Dim strEndPart As String
    Dim intLengthOfNewString As Integer
    
    'variables
    r = ""
    intLengthOfPieceToChop = Len(strPieceToChop)
    strEndPart = Right$(strContent, intLengthOfPieceToChop)
    intLengthOfNewString = Len(strContent) - intLengthOfPieceToChop
    
    'chop the text if it is at the end
    If strEndPart = strPieceToChop Then
        r = Mid$(strContent, 1, intLengthOfNewString)
    Else
        r = strContent
    End If
    
    qstr_ChopRight = r

End Function

Public Function qstr_PadTwoZeros(strLine As String) As String
    qstr_PadTwoZeros = Format$(strLine, "00")
End Function

Function qstr_FindFromRight(strFind As String, strMain As String) As Long
    Dim i As Integer
    
    For i = Len(strMain) To 1 Step -1
        If InStr(i, strMain, strFind) > 0 Then
            qstr_FindFromRight = i
            Exit Function
        End If
    Next i
    qstr_FindFromRight = 0
End Function

Public Function qstrEndsWith(str As String, ending As String) As Boolean
     Dim endingLen As Integer
     endingLen = Len(ending)
     EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function
 
Public Function qstrStartsWith(str As String, start As String) As Boolean
     Dim startLen As Integer
     startLen = Len(start)
     StartsWith = (Left(Trim(UCase(str)), startLen) = UCase(start))
End Function

' 343,2 --> 343.2
Function qstr_ForceFormatPeriodInDecimal(decValue As Variant)
    Dim strValue As String
    
    strValue = CStr(decValue)
    strValue = Replace(strValue, ",", ".")
    
    qstr_ForceFormatPeriodInDecimal = strValue
End Function
