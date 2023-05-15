Attribute VB_Name = "ArrayModule"
Option Explicit
'This Code Developed by Chris Vann Just copy it into a separate module, then call it from anywhere in your program as SortArray ArrayGoesHere

Public Function ArrayVuoto(myArray() As Variant) As Boolean
    ' Restituisce False se l'array è vuoto
    On Error Resume Next
    Dim ret As Long
    
    ret = 0
    
    If UBound(myArray) < 0 Then ret = Err.Number
    
    Select Case ret
        Case Is = 9 'Array is Nothing
            ArrayVuoto = True
        Case Is = 0 'Array is Nothing
            ArrayVuoto = False
        Case Else
            ArrayVuoto = True
    End Select

End Function

Public Function IsArrayVuoto(myArray) As Boolean
    ' Restituisce False se l'array è vuoto
    On Error Resume Next
    Dim ret As Long
    
    ret = 0
    
    If UBound(myArray) < 0 Then ret = Err.Number
    
    Select Case ret
        Case Is = 9 'Array is Nothing
            IsArrayVuoto = True
        Case Is = 0 'Array is Nothing
            IsArrayVuoto = False
        Case Else
            IsArrayVuoto = True
    End Select

End Function

Public Sub ArrayRemoveItem(ByRef ItemArray As Variant, ByVal ItemElement As Long)
    
    'PURPOSE:       Remove an item from an array, then
    '               resize the array
    
    'PARAMETERS:    ItemArray: Array, passed by reference, with
    '               item to be removed.  Array must not be fixed
    
    '               ItemElement: Element to Remove
                    
    'EXAMPLE:
    '           dim iCtr as integer
    '           Dim sTest() As String
    '           ReDim sTest(2) As String
    '           sTest(0) = "Hello"
    '           sTest(1) = "World"
    '           sTest(2) = "!"
    '           ArrayRemoveItem sTest, 1
    '           for iCtr = 0 to ubound(sTest)
    '               Debug.print sTest(ictr)
    '           next
    '
    '           Prints
    '
    '           "Hello"
    '           "!"
    '           To the Debug Window
    
    Dim lCtr As Long
    Dim lTop As Long
    Dim lBottom As Long
    
    If Not IsArray(ItemArray) Then
        Err.Raise 13, , "Type Mismatch"
        Exit Sub
    End If
    
    lTop = UBound(ItemArray)
    lBottom = LBound(ItemArray)
    
    If ItemElement < lBottom Or ItemElement > lTop Then
        Err.Raise 9, , "Subscript out of Range"
        Exit Sub
    End If
    
    For lCtr = ItemElement To lTop - 1
        ItemArray(lCtr) = ItemArray(lCtr + 1)
    Next
    
    On Error GoTo ErrorHandler:
    
    ReDim Preserve ItemArray(lBottom To lTop - 1)
    
    Exit Sub
    
ErrorHandler:
      'An error will occur if array is fixed
      Err.Raise Err.Number, , "You must pass a resizable array to this function"
        
End Sub

Public Function FilterDuplicates(arr As Variant) As Long
    ' Filter out duplicate values in an array and compact
    ' the array by moving items to "fill the gaps".
    ' Returns the number of duplicate values
    '
    ' it works with arrays of any type, except objects
    '
    ' The array is not REDIMed, but you can do it easily using
    ' the following code:
    '     a() is a string array
    '     dups = FilterDuplicates(a())
    '     If dups Then
    '         ReDim Preserve a(LBound(a) To UBound(a) - dups) As String
    '     End If
    '
    Dim col As Collection, index As Long, dups As Long
    Set col = New Collection
    
    On Error Resume Next
    
    For index = LBound(arr) To UBound(arr)
        ' build the key using the array element
        ' an error occurs if the key already exists
        col.Add 0, CStr(arr(index))
        If Err Then
            ' we've found a duplicate
            arr(index) = Empty
            dups = dups + 1
            Err.Clear
        ElseIf dups Then
            ' if we've found one or more duplicates so far
            ' we need to move elements towards lower indices
            arr(index - dups) = arr(index)
            arr(index) = Empty
        End If
    Next
    
    ' return the number of duplicates
    FilterDuplicates = dups
    
End Function
