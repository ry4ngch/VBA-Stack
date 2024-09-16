Attribute VB_Name = "Transpose2D"

Function Transpose2DArray(ByRef sourceArray() As Variant) As Variant()
    ' sourceArray() is type variant() means that it only accepts array as inputs.
    ' The sourceArray cannot be type variant
    ' To use the function, VBA sub or functions need to be declared as a variant() type to be accepted.
    
    ' Written by: Ryan Goh - 3 Oct 2023

    Dim currentRow As Long
    Dim LowerBoundRow As Long
    Dim UpperBoundRow As Long
    Dim currentColumn As Long
    Dim LowerBoundCol As Long
    Dim UpperBoundCol As Long
    Dim result() As Variant
    
    LowerBoundCol = LBound(sourceArray, 1)
    UpperBoundCol = UBound(sourceArray, 1)
    LowerBoundRow = LBound(sourceArray, 2)
    UpperBoundRow = UBound(sourceArray, 2)
   
    ReDim result(LowerBoundRow To UpperBoundRow, LowerBoundCol To UpperBoundCol)
    
    For currentRow = LowerBoundRow To UpperBoundRow
        For currentColumn = LowerBoundCol To UpperBoundCol
            result(currentRow, currentColumn) = sourceArray(currentColumn, currentRow)
        Next
    Next
    Transpose2DArray = result
End Function






