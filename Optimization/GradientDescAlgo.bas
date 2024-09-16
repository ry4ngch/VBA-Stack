Attribute VB_Name = "GradientDescAlgo"
'https://www.datasciencecentral.com/learn-under-the-hood-of-gradient-descent-algorithm-using-excel/

Function GradientDescentLinear(theta0 As Double, theta1 As Double, alpha As Double, iterations As Integer, X As Range, Y As Range, Optional returnParam As String = "0") As Variant
    'Here, theta0 and theta1 are the initial values of the parameters,
    'alpha is the learning rate, iterations is the number of iterations to run the gradient descent,
    'X is a range of the input values, and Y is a range of the target values.
    'The function calculates the cost function J and updates the values of theta0 and theta1 using the gradient descent algorithm.
    
    
    Dim m As Integer
    m = X.Rows.Count
    
    For i = 1 To iterations
        Dim j As Double
        j = 0
        
        Dim sum0 As Double
        sum0 = 0
        
        Dim sum1 As Double
        sum1 = 0
        
        For j = 1 To m
            sum0 = sum0 + (theta0 + theta1 * X(j, 1) - Y(j, 1))
            sum1 = sum1 + ((theta0 + theta1 * X(j, 1) - Y(j, 1)) * X(j, 1))
        Next j
        
        theta0 = theta0 - alpha * (1 / m) * sum0
        theta1 = theta1 - alpha * (1 / m) * sum1
    Next i
    
    If returnParam <> "0" Then
        If returnParam = "slope" Then
            GradientDescentLinear = theta1
        Else
            GradientDescentLinear = theta0
        End If
        Exit Function
    End If
    GradientDescentLinear = Array(theta0, theta1)
End Function

Public Function GradientDescentPoly(xtab As Variant, ytab As Variant, learningRate As Double, numIterations As Long, degree As Long) As Variant
    Dim m As Long
    If TypeName(xtab) = "Range" Then
        xtab = Application.Transpose(xtab.Value2)
        ytab = Application.Transpose(ytab.Value2)
    End If
    
    m = UBound(xtab)

    Dim n As Long
    n = degree + 1

    Dim theta() As Double
    ReDim theta(0 To n - 1)

    Dim X() As Double
    ReDim X(1 To m, 0 To n - 1)

    Dim i As Long, j As Long
    For i = 1 To m
        For j = 0 To n - 1
            X(i, j) = xtab(i) ^ j
        Next j
    Next i

    Dim h() As Double
    ReDim h(1 To m)

    Dim temp() As Double
    ReDim temp(0 To n - 1)

    Dim k As Long
    For k = 1 To numIterations
        For i = 1 To m
            h(i) = 0
            For j = 0 To n - 1
                h(i) = h(i) + theta(j) * X(i, j)
            Next j
        Next i

        For j = 0 To n - 1
            temp(j) = 0
            For i = 1 To m
                temp(j) = temp(j) + (h(i) - ytab(i)) * X(i, j)
            Next i
            temp(j) = theta(j) - learningRate * temp(j) / m
        Next j

        For j = 0 To n - 1
            theta(j) = temp(j)
        Next j
    Next k

    GradientDescentPoly = theta
End Function
