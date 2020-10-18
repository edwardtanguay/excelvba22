Option Explicit
 
Dim m_collection As ObjectArrayList
Dim m_searchText As String
 
Public Sub Constructor(Optional strSearchText As String = "")
    If strSearchText <> "" Then
        m_searchText = strSearchText
    End If
    FillCollection
End Sub
 
Private Sub Class_Initialize()
    m_searchText = ""
    Set m_collection = New ObjectArrayList
End Sub
 
Private Sub FillCollection()
 
    Dim objPerson As Person
    Dim intCount As Integer
    Dim x As Integer
    Dim row As Integer
    
    qexc_GotoWorksheet ("DataPersons")
 
    For row = 2 To 20
        
        Dim firstName As String
        Dim middleName As String
        Dim lastName As String
        Dim rank As Variant
        Dim searchString As String
        
        firstName = Trim(Cells(row, 1))
        middleName = Trim(Cells(row, 2))
        lastName = Trim(Cells(row, 3))
        rank = Cells(row, 4)
        searchString = UCase(firstName & "|" & middleName & "|" & lastName)
        
        If (searchString = "||") Then
            Exit For
        End If
 
        If InStr(searchString, UCase(m_searchText)) Then
            Set objPerson = New Person
            objPerson.firstName = firstName
            objPerson.middleName = middleName
            objPerson.lastName = lastName
            objPerson.rank = rank
        
            Call m_collection.Add(objPerson)
        End If
         
    Next
 
End Sub
 
Public Sub Display()
 
    Dim objPerson As Person
    Dim intIndex As Integer
 
    For intIndex = 0 To m_collection.NumberOfItems - 1
        Set objPerson = m_collection.GetItem(intIndex)
        Debug.Print objPerson.GetDisplayLine
    Next
 
End Sub
 
Public Sub CreateTextFile(pathAndFileName As String)
 
    Dim objPerson As Person
    Dim intIndex As Integer
    Dim content As String
    
    content = ""
 
    For intIndex = 0 To m_collection.NumberOfItems - 1
        Set objPerson = m_collection.GetItem(intIndex)
        content = content + objPerson.GetDisplayLine + vbNewLine
    Next
    
    qfil_CreateFile pathAndFileName, content
End Sub
