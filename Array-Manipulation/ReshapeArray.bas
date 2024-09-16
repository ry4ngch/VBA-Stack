Attribute VB_Name = "ReshapeArray"
Function ReshapeArray(arr As Variant, newRows As Long, newCols As Long)
    ' Reshape a 2-Dimensional array
    ' This ReshapeArray method works on the same concept as Python numpy reshape function
    ' Parameters:
    '   - arr: It is an array of vector or a matrix
    '   - newRows: number of rows for the new array to have
    '   - newCols: number of cols for the new array to have
    ' Written by: Ryan Goh Chuang Hong 10/04/2023
    
    Dim i As Long, j As Long
    Dim newArray() As Variant
    ReDim newArray(1 To newRows, 1 To newCols)
    
    If TypeName(arr) = "Range" Then
        arr = arr.Value2
    End If

    Dim counter As Long
    counter = 1
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            newArray((counter - 1) Mod newRows + 1, (counter - 1) \ newRows + 1) = arr(i, j)
            counter = counter + 1
        Next j
    Next i

    ReshapeArray = newArray
End Function
