'2 - PublicNotCreatable
Option Explicit
 
'internal variables
Dim m_arrContainer() As Variant
Dim m_intNumberOfItems As Integer
 
 
'method: add an item
Public Sub Add(varItem As Variant)
 
    'increment internal counter
    m_intNumberOfItems = m_intNumberOfItems + 1
 
    'redimension container
    ReDim Preserve m_arrContainer(m_intNumberOfItems)
 
    'now add the item
    Set m_arrContainer(m_intNumberOfItems - 1) = varItem
 
End Sub
 
'method: returns the number of items for for/next loops
Function NumberOfItems() As Integer
    NumberOfItems = m_intNumberOfItems
End Function
 
'method: return a specific item
Function GetItem(intIndexNumber As Integer) As Variant
    Set GetItem = m_arrContainer(intIndexNumber)
End Function
 
'method: returns whether or not item exists
Function ItemExists(varDesiredItem As Variant) As Boolean
 
    'declarations
    Dim intIndex As Integer
    Dim varItem As Variant
 
    'loop through and check
    For intIndex = 0 To m_intNumberOfItems - 1
 
        'variables
        varItem = Me.GetItem(intIndex)
 
        'if this is it, then return true
        If varItem = varDesiredItem Then
            ItemExists = True
            Exit Function
        End If
 
    Next
 
    'if we are here, it was not found
    ItemExists = False
 
End Function
