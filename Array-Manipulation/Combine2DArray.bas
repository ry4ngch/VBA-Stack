Attribute VB_Name = "Combine2DArray"

Function Combine2D(A As Variant, B As Variant, Optional stacked As Boolean = True) As Variant
    'assumes that A and B are 2-dimensional variant arrays
    'if stacked is true then A is placed on top of B
    'in this case the number of rows must be the same,
    'otherwise they are placed side by side A|B
    'in which case the number of columns are the same
    'LBound can be anything but is assumed to be
    'the same for A and B (in both dimensions)
    'False is returned if a clash

    ' Written by: Ryan Goh - 3 Oct 2023

    Dim lb As Long, m_A As Long, n_A As Long
    Dim m_B As Long, n_B As Long
    Dim m As Long, n As Long
    Dim i As Long, j As Long, k As Long
    Dim c As Variant

    If TypeName(A) = "Range" Then A = A.Value
    If TypeName(B) = "Range" Then B = B.Value

    lb = LBound(A, 1)
    m_A = UBound(A, 1)
    n_A = UBound(A, 2)
    m_B = UBound(B, 1)
    n_B = UBound(B, 2)

    If stacked Then
        m = m_A + m_B + 1 - lb
        n = n_A
        If n_B <> n Then
            Combine2D = False
            Exit Function
        End If
    Else
        m = m_A
        If m_B <> m Then
            Combine2D = False
            Exit Function
        End If
        n = n_A + n_B + 1 - lb
    End If
    ReDim c(lb To m, lb To n)
    For i = lb To m
        For j = lb To n
            If stacked Then
                If i <= m_A Then
                    c(i, j) = A(i, j)
                Else
                    c(i, j) = B(lb + i - m_A - 1, j)
                End If
            Else
                If j <= n_A Then
                    c(i, j) = A(i, j)
                Else
                    c(i, j) = B(i, lb + j - n_A - 1)
                End If
            End If
        Next j
    Next i
    Combine2D = c
End Function


