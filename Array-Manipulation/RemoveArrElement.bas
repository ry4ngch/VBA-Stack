Attribute VB_Name = "RemoveArrElement"
Function arrayPop(ByVal arr As Variant, ByVal indicesToRemove As Variant, ByVal byRow As Boolean) As Variant
    ''' Function Arguments:
    '       arr: type range or Variant() array
    '       indicesToRemove: can be integer, long, or variant() array. Note: the index starts from 0 for VBA Arrays, while index start from 1 for range
    '       byRow: set the arrayPop function to remove by row or column
    

    ''' Examples:
    '   1-D Horizontal Array as arr inputs:
    '   If the array is declared as array(1, 2, 3), the byRow argument should be set to False
    
    '   2-D Vertical Array as arr Inputs:
    '       Dim test(0 To 2, 0 To 0) As Variant
    '       test(0, 0) = 0
    '       test(1, 0) = 1
    '       test(2, 0) = 2
    '   To use the function, byRow should be pass as True
    
    ' For the function to work, getDims(), isInArray() and convertTo2D() function is also required
    ' Written by: Ryan Goh - 20/05/2023

    If TypeName(arr) = "Range" Then
        arr = arr.Value2
    End If
    
    Dim i As Long
    Dim index As Variant
    Dim k As Long
    Dim rEnd As Long
    Dim cEnd As Long
    Dim rStart As Long
    Dim cStart As Long
    
    If getDims(arr) = 1 Then ' means the array is 1-D horizontal
        arr = convertTo2D(arr)
    End If
    
    rEnd = UBound(arr, 1)
    cEnd = UBound(arr, 2)
    rStart = LBound(arr, 1)
    cStart = LBound(arr, 2)
    
    Dim j As Long: j = IIf(byRow, rStart, cStart)
    Dim tempArr() As Variant
    
    Dim removeCount As Long: removeCount = 1
    
    If TypeName(indicesToRemove) = "Variant()" Then
        removeCount = UBound(indicesToRemove) - LBound(indicesToRemove) + 1
    End If
    
    If byRow Then
        ReDim tempArr(rStart To rEnd - removeCount, cStart To cEnd)
    Else
        ReDim tempArr(rStart To rEnd, cStart To cEnd - removeCount)
    End If
    
    If TypeName(indicesToRemove) = "Double" Or TypeName(indicesToRemove) = "Integer" Or TypeName(indicesToRemove) = "Long" Then
        indicesToRemove = Array(indicesToRemove)
    End If

    If byRow Then
        For i = rStart To rEnd
            If Not IsInArray(i, indicesToRemove) Then
                For k = cStart To cEnd
                    tempArr(j, k) = arr(i, k)
                Next k
                j = j + 1
            End If
        Next i
    Else
        For k = cStart To cEnd
            If Not IsInArray(k, indicesToRemove) Then
                For i = rStart To rEnd
                    tempArr(i, j) = arr(i, k)
                Next i
                j = j + 1
            End If
        Next k
    End If
    
    arrayPop = tempArr
End Function
