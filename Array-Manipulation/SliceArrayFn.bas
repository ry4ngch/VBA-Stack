Attribute VB_Name = "SliceArrayFn"
Option Base 0

Function sliceArray(ByVal arr As Variant, _
                    ByVal startRow As Long, ByVal startColumn As Long, Optional ByVal rowStep As Long = 1, Optional ByVal colStep As Long = 1, _
                    Optional ByVal endRow As Long = -1, Optional ByVal endColumn As Long = -1) As Variant
    
    ' This function extracts a subarray from a given 2D array `arr` based on specified parameters.
   
    ' Parameters:
    '   - arr: The input 2D array.
    '   - startRow and startColumn: Starting indices for slicing.
    '   - rowStep and colStep (optional): Step sizes for row and column indices.
    '   - endRow and endColumn (optional): Ending indices for slicing. Default to the end of the array if not provided.

    ' Note: If endRow or endColumn are not specified (default values -1), they are set to the bounds of the input array.

    ' A new array `res` is created with dimensions based on the slice range and step sizes.
    ' The function iterates over the columns and rows of the slice area.
    ' For each position, it calculates the corresponding indices in the original array and populates the result array `res`.

    ' Transpose Check**:
    '   - The result array `res` is transposed for dimensional compatibility using `Application.Transpose`.
    '   - If the transposed result is effectively one-dimensional, it is converted to a specific type using `CVArrayBase`.

    ' The function returns the sliced subarray `res`, either as is or in a specific format depending on its dimensionality.

    ' Written by: Ryan Goh - 23 Oct 2023
    
    Dim i As Integer
    Dim j As Integer
    Dim ii As Integer
    Dim jj As Integer
    Dim cStep As Integer: cStep = 0
    Dim rStep As Integer: rStep = 0
    Dim testArr As Variant
    Dim res As Variant
    If endRow = -1 Then
        endRow = UBound(arr, 1)
    End If
    
    If endColumn = -1 Then
        endColumn = UBound(arr, 2)
    End If
    
    ReDim res(0 To (endRow - startRow) \ rowStep, 0 To (endColumn - startColumn) \ colStep)
    
    For j = 0 To endColumn - startColumn
        jj = j * cStep + startColumn
        If jj > endColumn Then Exit For
        For i = 0 To endRow - startRow
            ii = i * rStep + startRow
            If ii > endRow Then Exit For
            res(i, j) = arr(ii, jj)
            rStep = rowStep
        Next i
        cStep = colStep
    Next j
    
    testArr = Application.Transpose(res)
    If GetDims(testArr) = 1 Then
        sliceArray = CVArrayBase(testArr)
    Else
        sliceArray = res
    End If
End Function
