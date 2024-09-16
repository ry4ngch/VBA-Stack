Attribute VB_Name = "ShiftArrayBase"

Function ShiftArrayBase(ByVal sourceArray As Variant, Optional ArrayBase As Variant) As Variant
    ' This function shifts the base of 2D Array
    ' Inputs:
    '   sourceArray: can be type Variant() or range
    '   ArrayBase: by default the function will assign 0 if the ArrayBase is null. Meaning that the sourceArray base will shift from its original LBound to the assigned arrayBase
    ' Written by: Ryan Goh - 3 Oct 2023

    Dim currentRow As Long
    Dim currentCol As Long
    Dim result() As Variant
    If IsMissing(ArrayBase) Then ArrayBase = 0
    Dim rowShiftBase As Integer: rowShiftBase = LBound(sourceArray, 1) - ArrayBase
    Dim colShiftBase As Integer: colShiftBase = LBound(sourceArray, 2) - ArrayBase
    ArrayUBoundRow = UBound(sourceArray, 1)
    ArrayUBoundCol = UBound(sourceArray, 2)
    ReDim result(ArrayBase To ArrayUBoundRow - rowShiftBase, ArrayBase To ArrayUBoundCol - colShiftBase)
    
    For currentRow = ArrayBase To ArrayUBoundRow - rowShiftBase '[1 to 3] -> [0 to 2]
        For currentCol = ArrayBase To ArrayUBoundCol - colShiftBase
            result(currentRow, currentCol) = sourceArray(currentRow + rowShiftBase, currentCol + colShiftBase)
        Next currentCol
    Next currentRow
    ShiftArrayBase = result
End Function



