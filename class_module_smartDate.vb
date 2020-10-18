'USAGE EXAMPLE:
'------------------------
'Dim smartDate As smartDate
'Set smartDate = New smartDate
'monthStartDate = CDate("01.01." & CStr(calendarYear + 1))
'smartDate.Constructor monthStartDate
'smartDate.DisplayDebugInfo
'Sheets(wksMain).Cells(1, monthTitleColumn).Value = smartDate.theMonthName


Option Explicit

Dim m_theDate As Date
Dim m_theWeekdayNumber As Integer
Dim m_theWeekdayName As String
Dim m_theMonthNumber As Integer
Dim m_theMonthName As String
Dim m_theNumberOfDaysInMonth As Integer
Dim m_theWeekdayAbbreviation As String

'theDate
Property Get theDate() As Integer
    theDate = m_theDate
End Property
Property Let theDate(datTheDate As Integer)
    m_theDate = datTheDate
End Property

'theWeekdayNumber
Property Get theWeekdayNumber() As Integer
    theWeekdayNumber = m_theWeekdayNumber
End Property
Property Let theWeekdayNumber(intTheWeekdayNumber As Integer)
    m_theWeekdayNumber = intTheWeekdayNumber
End Property

'theWeekdayName
Property Get theWeekdayName() As String
    theWeekdayName = m_theWeekdayName
End Property
Property Let theWeekdayName(intTheWeekdayName As String)
    m_theWeekdayName = intTheWeekdayName
End Property

'theMonthNumber
Property Get theMonthNumber() As Integer
    theMonthNumber = m_theMonthNumber
End Property
Property Let theMonthNumber(intTheMonthNumber As Integer)
    m_theMonthNumber = intTheMonthNumber
End Property

'theMonthName
Property Get theMonthName() As String
    theMonthName = m_theMonthName
End Property
Property Let theMonthName(intTheMonthName As String)
    m_theMonthName = intTheMonthName
End Property

'theNumberOfDaysInMonth
Property Get theNumberOfDaysInMonth() As Integer
    theNumberOfDaysInMonth = m_theNumberOfDaysInMonth
End Property
Property Let theNumberOfDaysInMonth(intTheNumberOfDaysInMonth As Integer)
    m_theNumberOfDaysInMonth = intTheNumberOfDaysInMonth
End Property

'theWeekdayAbbreviation
Property Get theWeekdayAbbreviation() As String
    theWeekdayAbbreviation = m_theWeekdayAbbreviation
End Property
Property Let theWeekdayAbbreviation(intTheWeekdayAbbreviation As String)
    m_theWeekdayAbbreviation = intTheWeekdayAbbreviation
End Property

Public Function getBaseDate()
    getBaseDate = m_theDate
End Function

Public Sub Constructor(theDate As Date)

    'theDate
    m_theDate = theDate

    'theWeekday
    m_theWeekdayNumber = weekday(theDate) - 1
    If m_theWeekdayNumber = 0 Then
        m_theWeekdayNumber = 7
    End If
    
    'theWeekdayName
    m_theWeekdayName = weekdayName(m_theWeekdayNumber)
    
    'theMonthNumber
    m_theMonthNumber = Month(m_theDate)
    
    'theMonthName
    m_theMonthName = MonthName(m_theMonthNumber)
    
    'theNumberOfDaysInMonth
    m_theNumberOfDaysInMonth = qdat_numberOfDaysInMonth(m_theDate)
    
    'theWeekdayAbbreviation
    m_theWeekdayAbbreviation = Left(m_theWeekdayName, 2)
    
End Sub

Public Sub DisplayDebugInfo()
    Debug.Print "DATE: " & m_theDate
    Debug.Print "theWeekdayNumber: " & m_theWeekdayNumber
    Debug.Print "theWeekdayName: " & m_theWeekdayName
    Debug.Print "theMonthNumber: " & m_theMonthNumber
    Debug.Print "theMonthName: " & m_theMonthName
    Debug.Print "theNumberOfDaysInMonth: " & m_theNumberOfDaysInMonth
    Debug.Print "theWeekdayAbbreviation: " & m_theWeekdayAbbreviation
    Debug.Print "---"
    
End Sub
