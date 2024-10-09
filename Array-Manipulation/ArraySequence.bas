Attribute VB_Name = "ArraySequence"
Function ArrSeq(lRows As Long, Optional lCols As Long = 1, Optional lStart As Double = 1, Optional lStep As Double = 1) As Variant
    ' This function works similar to Excel Sequence() in-built function.
    ' The purpose why this function was developed is because older excel version does not have the in-built Sequence()
    
    ' Arguments:
    '   - lRows - number of rows for the sequence
    '   - lCols - number of column for the sequence
    '   - lStart - starting number for the sequence
    '   - lStep - step size for the sequence.
    
    ' Example usage:
    '   ArrSeq(5, 3, 1, 2)
    '   This will produce 5 rows and 3 columns staring with 1 and step size of 2.

  Dim r As Long, c As Long, Nums As Double
  Nums = lStart
  For r = 1 To lRows
    For c = 1 To lCols
      ArrSeq = ArrSeq & "," & Nums
      Nums = Nums + lStep
    Next
    ArrSeq = ArrSeq & ";"
  Next
  ArrSeq = Evaluate("{" & Replace(Mid(Left(ArrSeq, Len(ArrSeq) - 1), 2), ";,", ";") & "}")
End Function
