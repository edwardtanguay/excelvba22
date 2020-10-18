 Option Explicit
  
 Dim m_firstName As String
 Dim m_middleName As String
 Dim m_lastName As String
 Dim m_rank As Variant
  
 'FirstName
 Property Get firstName() As String
     firstName = m_firstName
 End Property
 Property Let firstName(strFirstName As String)
     m_firstName = strFirstName
 End Property
  
 'MiddleName
 Property Get middleName() As String
     middleName = m_middleName
 End Property
 Property Let middleName(strMiddleName As String)
     m_middleName = strMiddleName
 End Property
  
 'LastName
 Property Get lastName() As String
     lastName = m_lastName
 End Property
 Property Let lastName(strLastName As String)
     m_lastName = strLastName
 End Property
  
 'Rank
 Property Get rank() As Variant
     rank = m_rank
 End Property
 Property Let rank(decRank As Variant)
     m_rank = decRank
 End Property
  
 Public Sub Constructor(strFirstName As String, strLastName As String)
     m_firstName = strFirstName
     m_lastName = strLastName
 End Sub
  
 Private Sub Class_Initialize()
     m_firstName = ""
     m_middleName = ""
     m_lastName = ""
     m_rank = CDec(0)
 End Sub
 
 Function GetDisplayLine()
     Dim strRank As String
     Dim strSmartMiddleName As String
  
     strRank = qstr_ForceFormatPeriodInDecimal(m_rank)
     If m_middleName = "" Then
         strSmartMiddleName = ""
     Else
         strSmartMiddleName = " " & m_middleName
     End If
     GetDisplayLine = m_firstName & strSmartMiddleName & " " & m_lastName & " (" & strRank & ")"
 End Function

