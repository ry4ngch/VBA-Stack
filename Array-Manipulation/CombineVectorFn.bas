Attribute VB_Name = "CombineVectorFn"

Function CombineVector(ByVal stackedAxis As Long, ParamArray arrays() As Variant) As Variant
    ' The `CombineVector` function combines multiple arrays (vectors) into a single array based on the specified axis for stacking.

    ' Parameters:
    '   - stackedAxis: Indicates how the arrays should be combined (0 for vertical stacking, 1 for horizontal stacking).
    '   - arrays(): A variable number of input arrays to be combined.

    ' Vertical Stacking (stackedAxis = 0):
    '   - Computes the total number of rows by summing up the sizes of all input arrays.
    '   - Creates a single-column array `arr` with the total row count.
    '   - Iterates through each input array and appends its elements to `arr` in sequence.

    ' Horizontal Stacking (stackedAxis = 1):
    '   - Determines the maximum number of rows needed by checking the size of each input array.
    '   - Creates a 2D array `arr` with the maximum row count and the number of columns equal to the number of input arrays.
    '   - Fills in `arr` by placing elements from each input array into the corresponding column.

    '  The function returns the combined array `arr`, either as a single column vector or a 2D matrix, depending on the stacking axis.

    ' `CombineVector` combines multiple arrays into a single array either by stacking them vertically or horizontally, depending on the `stackedAxis` parameter.

    ' Written by: Ryan Goh - 23 Oct 2023

    Dim i As Long
    Dim j As Long
    Dim k As Long: k = 0
    Dim n_rows As Long
    Dim n_cols As Long
    Dim arr() As Variant

    n_cols = UBound(arrays) 'number of vectors
    
    If stackedAxis = 0 Then
        n_rows = UBound(arrays)
        For j = 0 To n_cols
            n_rows = n_rows + UBound(arrays(j))
        Next j
        
        ReDim arr(0 To n_rows)
        For j = 0 To n_cols
            For i = 0 To n_rows - UBound(arrays(j)) - UBound(arrays)
                arr(k) = arrays(j)(i)
                k = k + 1
            Next i
        Next j
        
    Else
        n_rows = UBound(arrays(0))
        For j = 0 To n_cols
            If UBound(arrays(j)) > n_rows Then
                n_rows = UBound(arrays(j))
            End If
        Next j

        ReDim arr(0 To n_rows, 0 To n_cols)
        For i = 0 To n_rows
            For j = 0 To n_cols
                arr(i, j) = arrays(j)(i)
            Next j
        Next i
    End If

    CombineVector = arr
    
End Function
