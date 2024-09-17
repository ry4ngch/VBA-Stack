Attribute VB_Name = "NewtonRaphson"
Function NewtonRaphsonOptimization(func As Variant, dfunc As Variant, guess As Double, tolerance As Double, max_iteration As Integer) As Double
    ' Parameters:
    '   - func: Accept `func` as a string
    '   - dfunc: Accept `dfunc` as a string, make sure dfunc is the derivative of the equation defined in `func`
    '   - guess: make an initial guess for the algorithm to start
    '   - tolerance: set the convergence criteria for the solution
    '   - max_iteration: set the max iteration for the loop
    
    Dim x As Double
    Dim fx As Double
    Dim dfx As Double
    Dim i As Integer
    
    ' Initialize variables
    x = guess
    i = 0
    
    ' Iterate to find the root
    Do While i < max_iteration
        fx = Application.Run(func, x)
        dfx = Application.Run(dfunc, x)
        
        ' Check if the derivative is zero (avoid division by zero)
        If dfx = 0 Then
            ' Derivative is zero. No solution found.
            NewtonRaphsonOptimization = CVErr(xlErrDiv0)
            Exit Function
        End If
        
        ' Update x using the Newton-Raphson formula
        x = x - fx / dfx
        
        ' Check for convergence
        If Abs(fx) < tolerance Then
            NewtonRaphsonOptimization = x
            Exit Function
        End If
        
        ' Increment iteration counter
        i = i + 1
    Loop
    
    ' If max iterations reached, return the last computed value
    NewtonRaphsonOptimization = x
End Function

' For testing purposes
' Define the function to optimize
Function f(x As Double) As Double
    ' Example function: f(x) = x^2 - 4
    f = x ^ 2 - 4
End Function

' Define the derivative of the function
Function df(x As Double) As Double
    ' Example derivative: df/dx = 2*x
    df = 2 * x
End Function

Sub TestNewtonRaph()
    Dim result As Double
    Dim tol As Double: tol = 0.0001
    Dim max_iter As Integer: max_iter = 100
    Dim guess As Double: guess = 1

    ' Call optimization function
    result = NewtonRaphsonOptimization("f", "df", guess, tol, max_iter)
    Debug.Print "The result of the Newton-Raphson Optimization is: " & result
End Sub
