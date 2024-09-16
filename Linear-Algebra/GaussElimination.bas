Attribute VB_Name = "GaussElimination"

Function GaussSolver(ByVal A As Variant, ByVal B As Variant) As Variant
    ''' This function solves system of equations and returns the solution vector
    ''' The function is used together with ShiftArrayBase and CVArrayBase which converts the matrix and vector to base 0
    '   This is done so that the function can be used both in Excel as a user defined function or as a self contained module
    '   in other code blocks, streamlining the programming complexity

    ''' The code for gaussian elimination with back substitution perform the following task:
    '   1. Forward Elimination: Converts the system of linear equations into an upper triangular matrix by eliminating
    '   variables from below each pivot position through row operation.
    '   2. Back substitution: Solves for the variables starting from the last row of the upper triangular matrix and working
    '   upwards, substituting known values to find the remaining variables.

    ''' Inputs:
    '   A - Coefficients of equation (2D matrix)
    '   B - 1D Vectors (constants term of the system)

    ''' Outputs:
    ' The function returns a vector consisting of all possible roots of the system of linear equation.

    ' Written by: Ryan Goh - 3 Oct 2023

    Dim n As Long
    
    If TypeName(A) = "Range" Then
        Dim OVb As Range: Set OVb = B
        A = ShiftArrayBase(A.Value2, 0)
        B = CVArrayBase(Application.Transpose(B.Value2))
    End If
    
    n = UBound(A) 'row of matrices
    
    Dim i As Long, j As Long, k As Long, maxIdx As Long
    Dim tempA As Double, tempB As Double, sum As Double

    For k = 0 To n - 1
        maxIdx = k
        
        ' Identify which row have the largest element
        For i = k + 1 To n
            If Abs(A(i, k)) > Abs(A(maxIdx, k)) Then maxIdx = i
        Next i
        
        ' Swap Rows
        If maxIdx <> k Then
            For j = 0 To n
                tempA = A(k, j)
                A(k, j) = A(maxIdx, j)
                A(maxIdx, j) = tempA
            Next j
            tempB = B(k)
            B(k) = B(maxIdx)
            B(maxIdx) = tempB
        End If
        
        ' Forward Elimination
        For i = k + 1 To n
            factor = A(i, k) / A(k, k)
            For j = k + 1 To n
                A(i, j) = A(i, j) - factor * A(k, j)
            Next j
            B(i) = B(i) - factor * B(k)
        Next i
    Next k
    

    ' Backward Substitution
    Dim X() As Double
    ReDim X(0 To n)
    
    ' Get the x values of the last row
    X(n) = B(n) / A(n, n)
    
    ' Get the x values of the other rows except for the last row
    For i = n - 1 To 0 Step -1
        sum = 0
        For j = i + 1 To n
            sum = sum + A(i, j) * X(j)
        Next j
        X(i) = (B(i) - sum) / A(i, i)
    Next i
    
    
    If TypeName(OVb) = "Range" Then
        If (OVb.Columns.Count = 1) Then
            GaussSolver = Application.Transpose(X)
        Else
            GaussSolver = X
        End If
    Else
        GaussSolver = X
    End If
End Function

