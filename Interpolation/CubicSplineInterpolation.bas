Attribute VB_Name = "CubicSplineInterpolation"
Function Cubic(known_x As Variant, known_y As Variant, x_interp As Variant, Optional outputType As String = "yhat") As Variant
    ' Conditions:
    '   - Condition 1: Each interior nodes need to have 2 equations
    '   - Condition 2: The 1st and and last nodes must pass through the end points (2 equations)
    '   - Condition 3: The 1st derivative at the interior nodes must be equal (ie: between 2 neighbouring points)
    '   - Condition 4: The 2nd derivative at the interior nodes must be equal (ie: between 2 neighbouring points)
    '   - Condition 5: the 2nd derivative at the end nodes (1st and last point) is 0
    
    ' Example:
    ' For eg: For a 10 data points (we will need 36 equations for 36 unknowns -> 9 intervals x 4 unknowns for each equation)
    '   - Condition 1: 8 in-between knots each need to have 2 equations => 16 equations
    '   - Condition 2: The last 2 nodes (1 at the beginning and 1 at the last point) each have 1 equations => 2 equations
    '   - Condition 3: the 1st derivative of the 8 interior nodes -> creates 8 equations
    '   - Condition 4: the 2nd derivative of the 8 interior nodes -> creates 8 equations
    '   - Condition 5: the 2nd derivative at the end nodes (1st and last point) is 0 -> creates 2 equations
    
    
    
    Dim X As Variant
    Dim Y As Variant
    Dim x_u As Variant
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '2 - Adjust input locations
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Input x and y values
    If TypeName(known_x) = "Range" Then
        X = CVArrayBase(Application.Transpose(known_x.Value2))
        Y = CVArrayBase(Application.Transpose(known_y.Value2))
    Else
        X = known_x
        Y = known_y
    End If
    
    'Input xu values to be interpolated
    If TypeName(x_interp) = "Range" Then
        If x_interp.Count > 1 Then
            x_u = CVArrayBase(Application.Transpose(x_interp.Value2))
        Else
            x_u = Array(x_interp.Value2)
        End If
    ElseIf TypeName(x_interp) = "Double" Or TypeName(x_interp) = "Long" Or TypeName(x_interp) = "Integer" Then
        x_u = Array(x_interp)
    Else
        x_u = x_interp
    End If
    
    'Number of Data Points to create the interpolation function
    n = UBound(X)
    'Number of Data point to be interpolated (Could include the values in the original data-set)
    m = UBound(x_u)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ReDim e(n)
    ReDim g(n)
    ReDim r(n)
    ReDim f(n)
    ReDim d2x(n) '2nd derivatives
    ReDim factor(n)
    ReDim output(m) As Variant
    ReDim derivative1(m) As Variant
    ReDim derivative2(m) As Variant
    
    '1st point and last point 2nd derivative is zero -> condition 5
    d2x(0) = 0
    d2x(n) = 0

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Create Tridiagonal Matrix
    '   e2 to e8
    '   f1 to f8
    '   g1 to g7

    For i = 1 To n - 1
        If i > 1 Then
            e(i) = (X(i) - X(i - 1))
        End If
        f(i) = 2 * (X(i + 1) - X(i - 1))
        If i < n - 1 Then
            g(i) = (X(i + 1) - X(i))
        End If
        
        r(i) = (6 / (X(i + 1) - X(i))) * (Y(i + 1) - Y(i))
        r(i) = r(i) + (6 / (X(i) - X(i - 1))) * (Y(i - 1) - Y(i))

    Next i
    

    'Thomas Algorithm
    'Forward Elimination
    For k = 2 To n - 1
        factor(k) = e(k) / f(k - 1)
        f(k) = f(k) - factor(k) * g(k - 1)
        r(k) = r(k) - factor(k) * r(k - 1)
    Next k

    'Backward Substitution
    d2x(n - 1) = r(n - 1) / f(n - 1)
    For k = n - 2 To 1 Step -1
        d2x(k) = (r(k) - g(k) * d2x(k + 1)) / f(k)
    Next k
        
    'Interpolation
    For j = 0 To m
        xu = x_u(j)
        flag = 0
        i = 1
        Do
            'Identify what interval the pt of interest is
            If xu >= X(i - 1) And xu <= X(i) Then
                'Calculating Constants
        
                c1 = d2x(i - 1) / (6 * (X(i) - X(i - 1)))
                c2 = d2x(i) / (6 * (X(i) - X(i - 1)))
                c3 = Y(i - 1) / (X(i) - X(i - 1)) - d2x(i - 1) * (X(i) - X(i - 1)) / 6
                c4 = Y(i) / (X(i) - X(i - 1)) - d2x(i) * (X(i) - X(i - 1)) / 6
        
                'Function Value at xu
                t1 = c1 * (X(i) - xu) ^ 3
                t2 = c2 * (xu - X(i - 1)) ^ 3
                t3 = c3 * (X(i) - xu)
                t4 = c4 * (xu - X(i - 1))
                yu = t1 + t2 + t3 + t4
        
                '1st Derivative at xu
                t1 = -3 * c1 * (X(i) - xu) ^ 2
                t2 = 3 * c2 * (xu - X(i - 1)) ^ 2
                t3 = -c3
                t4 = c4
                dy = t1 + t2 + t3 + t4
        
                '2nd Derivative at xu
                t1 = 6 * c1 * (X(i) - xu)
                t2 = 6 * c2 * (xu - X(i - 1))
                d2y = t1 + t2
                flag = 1
            Else
                i = i + 1
            End If
            If i = n + 1 Or flag = 1 Then
                Exit Do
            End If
        Loop
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' 3 - Define where you want the interpolated value and derivatives to be outputted
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        derivative1(j) = dy
        derivative2(j) = d2y
        output(j) = yu

    Next j
    
    If TypeName(x_interp) = "Range" Then
        If (x_interp.Columns.Count = 1) Then
            output = Application.Transpose(output)
            derivative1 = Application.Transpose(derivative1)
            derivative2 = Application.Transpose(derivative2)
        End If
    End If
    
    If outputType = "dy" Then
        Cubic = derivative1
    ElseIf outputType = "d2y" Then
        Cubic = derivative2
    Else
        Cubic = output
    End If
End Function

