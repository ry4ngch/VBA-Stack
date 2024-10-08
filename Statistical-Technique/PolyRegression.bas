Attribute VB_Name = "PolyRegression"
Option Base 0

Function PolynomialRegression(ByVal xData As Variant, ByVal yData As Variant, ByVal Order As Integer, Optional ByVal outputType As String = "coeff")
    ' This function solves for the coefficients of an equations which are generated by fitting to a number of datapoints by performing polynomial regressions.
    ' The function accepts an outputType below:
    ' - coeff: returns the coefficients of the polynomial regression model
    ' - r2: returns the r^2 score of the polynomial regression model

    ' This function is quite similar to Excel in built LINEST function
    ' The difference here is this function accepts both range or variant() type data and can be used directly in other VB Projects
    ' without referencing the worksheet function which tends to slow down excel.
    ' Also, the data can be passed in directly as an array within VBA IDE.

    ' Written by: Ryan Goh - 23 Oct 2023

    Dim X As Variant
    Dim Y As Variant
    If TypeName(xData) = "Range" Then
        X = CVArrayBase(Application.Transpose(xData.Value2))
        Y = CVArrayBase(Application.Transpose(yData.Value2))
    Else
        X = xData
        Y = yData
    End If
    
    Dim A() As Variant ' matrixA
    Dim B() As Variant ' vector B
    Dim s As Variant ' coefficients (solutions)
    ReDim A(0 To Order, 0 To Order)
    ReDim B(0 To Order)
    
    Dim n As Integer: n = UBound(X)
    
    ' Generate the matrixA and vector B for the inputted polynomial order
    For i = 0 To Order
        'Create A matrix
        For j = 0 To i
            k = i + j
            sum = 0
            For L = 0 To n
                sum = sum + X(L) ^ k
            Next L
            A(i, j) = sum
            A(j, i) = sum
        Next j
        
        'Create b vector
        sum = 0
        For L = 0 To n
            sum = sum + Y(L) * X(L) ^ i
        Next L
        B(i) = sum
    Next i
    
    
    ' Solve system of equation using gaussian elimination
    s = GaussSolver(A, B)
    
    
    ' Calculate r2
    ' Calculate St & Sr
    ym = B(0) / (n + 1) 'note: we use n+1 here since indices start from 0 base
    
    st = 0
    For i = 0 To n
        st = st + (Y(i) - ym) ^ 2
    Next i
    
    sr = 0
    For i = 0 To n
        sum = Y(i)
        For j = 0 To Order
            sum = sum - s(j) * X(i) ^ j
        Next j
        
        sr = sr + sum ^ 2
    Next i
    r2 = (st - sr) / st
    
    ' output the calculation (coefficients and r2 score)
    
    
    If TypeName(xData) = "Range" Then
        If (xData.Columns.Count = 1) Then
            s = Application.Transpose(s)
        End If
    End If
    
    If outputType = "coeff" Then
        PolynomialRegression = s
    Else
        PolynomialRegression = r2
    End If
    
    
End Function