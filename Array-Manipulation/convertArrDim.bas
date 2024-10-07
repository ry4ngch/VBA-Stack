Attribute VB_Name = "convertArrDim"
Function convertTo2D(ByVal arr As Variant, Optional ByVal baseOption As Long = 0) As Variant
    ' This function converts a 1-Dimensional array to 2-Dimensional
    ' Writen by: Ryan Goh - 23/05/2023
    
    ' Argument:
    '   - arr: 1D array
    '   - baseOption: optional inputs to convert the array to other starting base, default to base 0
    
    If getDims(arr) > 1 Then
        Exit Function
    End If
    
    Dim i As Long
    
    Dim tempArr() As Variant
    ReDim tempArr(baseOption To baseOption, baseOption To UBound(arr) + baseOption)

    For i = LBound(arr) To UBound(arr)
        tempArr(baseOption, i + baseOption) = arr(i)
    Next i

    convertTo2D = tempArr
    
End Function
