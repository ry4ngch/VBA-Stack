Attribute VB_Name = "MatrixSort"

Public Sub QuickSortMatrix(ByRef SortArray As Variant, Optional ByVal lngMin As Long = -1, Optional ByVal lngMax As Long = -1, Optional ByVal lngColumn As Long = 0, Optional ByVal bAscending As Boolean = True)
    On Error Resume Next

    'Sort a 2-Dimensional array
    'Posted by Jim Rech 10/20/98 Excel.Programming
    'Modifications, Nigel Heffernan:
    '       ' Escape failed comparison with empty variant
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If IsEmpty(SortArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortArray, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray, 1)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a valid default member or property
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
            While SortArray(i, lngColumn) < varMid And i < lngMax
                i = i + 1
            Wend
            While varMid < SortArray(j, lngColumn) And j > lngMin
                j = j - 1
            Wend
    
            If i <= j Then
                ' Swap the rows
                For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
                    arrRowTemp = SortArray(i, lngColTemp)
                    SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
                    SortArray(j, lngColTemp) = arrRowTemp
                Next lngColTemp
    
                i = i + 1
                j = j - 1
            End If
        Wend
    Else
        While (i <= j)
            While (SortArray(i, col) > varMid And i < lngMax)
                i = i + 1
            Wend
            
            While (varMid > SortArray(j, col) And j > lngMin)
                j = j - 1
            Wend
            
            If (i <= j) Then
                ' Swap the rows
                For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
                    arrRowTemp = SortArray(i, lngColTemp)
                    SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
                    SortArray(j, lngColTemp) = arrRowTemp
                Next lngColTemp
                i = i + 1
                j = j - 1
            End If
        Wend
    End If

    If (lngMin < j) Then Call QuickSortMatrix(SortArray, lngMin, j, lngColumn, bAscending)
    If (i < lngMax) Then Call QuickSortMatrix(SortArray, i, lngMax, lngColumn, bAscending)
    
End Sub

Function QuickSortedMatrix(ByVal arrayData As Variant, Optional ByVal lngMin As Long = -1, Optional ByVal lngMax As Long = -1, Optional ByVal lngCol As Long = 0, Optional ByVal bAscending As Boolean = True) As Variant
    ' This function is an extension of the QuickSortMatrix Subroutine by Jim Rech.
    ' The function extends the functionality of the subroutine by allowing it to be used as a UDF in excel
    ' It also provides a return matrix of the sorting matrix, which is useful for usage in other code blocks
    ' where matrix or array manipulation and interpolation is required 

    ' Written by: Ryan Goh - 23 Oct 2023
   
    Dim result As Variant
    If TypeName(arrayData) = "Range" Then
        result = ShiftArrayBase(arrayData.Value2, 0)
    Else
        result = arrayData
    End If
    
    QuickSortMatrix result, lngMin, lngMax, lngCol, bAscending
    QuickSortedMatrix = result
End Function

