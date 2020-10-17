' e.g. 12345 -> 15
Function qmat_SumOfDigits(num As Double) As Long
    Dim total As Long
    
    total = 0
    
    For i = 1 To Len(CStr(num))
        total = total + Mid(num, i, 1)
    Next i
    
    qmat_SumOfDigits = total
End Function