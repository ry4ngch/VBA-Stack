Attribute VB_Name = "LUDecomposition"
Function LUDecompose(A As Variant, B As Variant, Optional output As String) As Variant
    ' LUDecompose function that takes in a square matrix A and a vector B and performs LU decomposition.
    ' It returns the solution vector X.

    ''' Similar to the gauss elimination technique, this function is used together with ShiftArrayBase and CVArrayBase which converts the matrix and vector to base 0
    '   This is done so that the function can be used both in Excel as a user defined function or as a self contained module
    '   in other code blocks, streamlining the programming complexity

    ' Written by: Ryan Goh - 23 Oct 2023
    
    If TypeName(A) = "Range" Then
        Dim BRng As Range: Set BRng = B
        A = ShiftArrayBase(A.Value2, 0)
        B = CVArrayBase(Application.Transpose(B.Value2))
    End If

    Dim n As Integer, i As Integer, j As Integer, k As Integer
    Dim L() As Double, U() As Double, P() As Integer, X() As Double, LU() As Double
    Dim maxVal As Double, tmp As Double, mult As Double, pivot As Integer

    n = UBound(A, 1)                ' Get the size of the square matrix A.
    ReDim L(0 To n, 0 To n)         ' Create an array to hold the lower triangular matrix L.
    ReDim U(0 To n, 0 To n)         ' Create an array to hold the upper triangular matrix U.
    ReDim P(0 To n)                  ' Create an array to hold the permutation matrix P.
    ReDim X(0 To n)                  ' Create an array to hold the solution vector X.
    ReDim LU(0 To n, 0 To n)

    For i = 0 To n
        P(i) = i                     ' Initialize the permutation matrix to the identity matrix.
        For j = 0 To n
            If i = j Then
                L(i, j) = 1         ' Initialize the diagonal of L to 1.
            End If
            U(i, j) = A(i, j)       ' Initialize the U matrix to the input matrix A.
            LU(i, j) = A(i, j)      ' Initialize the LU matrix to the input matrix A. Note: this LU matrix is not neccessary for computing the system of equations. It is more for outputing purpose
        Next j
    Next i

    ' Perform LU decomposition with partial pivoting.
    For k = 0 To n - 1
        maxVal = 0
        For i = k To n
            If Abs(U(i, k)) > maxVal Then
                maxVal = Abs(U(i, k))
                pivot = i
            End If
        Next i

        If maxVal = 0 Then
            MsgBox "Error: Matrix is singular"
            Exit Function
        End If

        If pivot <> k Then
            ' Swap rows k and tmp in U and L.
            For j = 0 To n
                ' Swap entries in U.
                tmp = U(k, j)
                U(k, j) = U(pivot, j)
                U(pivot, j) = tmp
                
                ' Swap entries in LU
                tmp = LU(k, j)
                LU(k, j) = LU(pivot, j)
                LU(pivot, j) = tmp
                
                ' Swap entries in L.
                tmp = L(k, j)
                L(k, j) = L(pivot, j)
                L(pivot, j) = tmp
            Next j

            ' Swap entries k and tmp in P.
            tmp = P(k)
            P(k) = P(pivot)
            P(pivot) = tmp
        End If

        ' Compute the entries of L and U.
        ' Forward Elimination
        For i = k + 1 To n
            mult = U(i, k) / U(k, k)
            multLU = LU(i, k) / LU(k, k)
            L(i, k) = mult
            LU(i, k) = multLU
            For j = k To n
                U(i, j) = U(i, j) - mult * U(k, j)
            Next j
            For j = k + 1 To n
                LU(i, j) = LU(i, j) - multLU * LU(k, j)
            Next j
        Next i
    Next k

    Select Case output
        Case "U"
            LUDecompose = U
        Case "P"
            LUDecompose = P
        Case "L"
            LUDecompose = L
        Case "LU"
            LUDecompose = LU
        Case Else
            Dim OrientVert As Boolean
            If Not BRng Is Nothing Then
                OrientVert = (BRng.Columns.Count = 1)
            Else
                OrientVert = False
            End If
            LUDecompose = LUSolve(L, U, P, B, OrientVert)
    End Select

End Function

Function LUSolve(L As Variant, U As Variant, P As Variant, B As Variant, OrientVert As Boolean) As Variant
    Dim n As Long
    Dim i As Long, j As Long, k As Long
    Dim Y() As Double
    Dim X() As Double
    Dim sum As Double
    
    ' Determine the size of the input matrices and vectors.
    n = UBound(L, 1)
    
    ' Initialize the Y and X vectors.
    ReDim Y(0 To n) As Double
    ReDim X(0 To n) As Double
    
    ' Apply the permutation matrix P to the input vector B.
    For i = 0 To n
        Y(i) = B(P(i))
    Next i
    
    ' Solve LY = B for Y using forward substitution.
    For i = 0 To n
        sum = 0
        For j = 0 To i - 1
            sum = sum + L(i, j) * Y(j)
        Next j
        Y(i) = Y(i) - sum
    Next i
    
    ' Solve UX = Y for X using backward substitution.
    For i = n To 0 Step -1
        sum = 0
        For j = i + 1 To n
            sum = sum + U(i, j) * X(j)
        Next j
        X(i) = (Y(i) - sum) / U(i, i)
    Next i
    
    ' Return the solution vector X.
    If OrientVert Then
        LUSolve = Application.Transpose(X)
    Else
        LUSolve = X
    End If
    
End Function

