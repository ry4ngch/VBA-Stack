Attribute VB_Name = "VectorSort"
Public Sub QuickSortVector(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional bAscending As Boolean = True)
    On Error Resume Next

    'Sort a 1-Dimensional array
    'Parameters:
    '   Mandatory Parameter
    '        - SortArray: must be vector input
    '   Optional Parameter
    '       - lngMin: Set the starting index for sorting range (For range type input, the starting index is 1, for array, starting index is 0 by default)
    '       - lngMax: set the ending index for sorting range (For range type input, the ending index is Range.count, for array , ending index is UBound(array) by default)
    '       - bAscending: set the output sorting to be either ascending or descending. By default, the procedure will sortby Ascending.
    ' Originally posted by Jim Rech 10/20/98 Excel.Programming
    ' Modifications, Nigel Heffernan:
    '       ' Escape failed comparison with an empty variant in the array
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim varX As Variant

    If IsEmpty(SortArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    
    ' If sorting range is not define, auto set sorting range to LBound and Ubound of the array
    If lngMin = -1 Then
        lngMin = LBound(SortArray)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2)

    ' We send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a default member or property
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    If bAscending Then
        While i <= j
            While SortArray(i) < varMid And i < lngMax
                i = i + 1
            Wend
            
            While varMid < SortArray(j) And j > lngMin
                j = j - 1
            Wend
    
            If i <= j Then
                ' Swap the item
                varX = SortArray(i)
                SortArray(i) = SortArray(j)
                SortArray(j) = varX
                i = i + 1
                j = j - 1
            End If
        Wend
    Else
        While (i <= j)
            While (SortArray(i) > varMid And i < lngMax)
                i = i + 1
            Wend

            While (varMid > SortArray(j) And j > lngMin)
                j = j - 1
            Wend
            If (i <= j) Then
                ' Swap the item
                varX = SortArray(i)
                SortArray(i) = SortArray(j)
                SortArray(j) = varX
                i = i + 1
                j = j - 1
            End If
        Wend
    End If

    If (lngMin < j) Then Call QuickSortVector(SortArray, lngMin, j, bAscending)
    If (i < lngMax) Then Call QuickSortVector(SortArray, i, lngMax, bAscending)

End Sub

Function QuickSortedVector(ByVal vector As Variant, Optional ByVal lngMin As Long = -1, Optional ByVal lngMax As Long = -1, Optional ByVal bAscending As Boolean = True) As Variant
    ' This function is a wrapper for QuickSortVector sub. 
    ' By using this function, results are return by the functions, which can be used directly in spreadsheet or return the output in code
    
    ' Written by: Ryan Goh - 23 Oct 2023
    
    Dim result As Variant
    If TypeName(vector) = "Range" Then
        result = Application.Transpose(vector.Value2)
    Else
        result = vector
    End If
    
    QuickSortVector result, lngMin, lngMax, bAscending
    QuickSortedVector = result
End Function

