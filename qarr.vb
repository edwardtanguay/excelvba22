Public Function qarr_IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    If stringToBeFound = "" Then
        qarr_IsInArray = False
        Exit Function
    End If
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            qarr_IsInArray = True
            Exit Function
        End If
    Next i
    qarr_IsInArray = False
End Function