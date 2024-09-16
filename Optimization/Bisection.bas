Attribute VB_Name = "Bisection"
Function func(ByVal X As Double)
    func = X ^ 2 - X - 2
End Function


Public Function BisectionMethod(ByVal f As String, ByVal xl As Double, ByVal xu As Double, ByVal tol As Double, Optional displayFullRes As Boolean = False) As Variant
    'The bisection method is used to find the roots of an equation
    'xl: lower x interval
    'xu: upper x interval
    'f: function passed as string
    'tol: tolerance to meet before while loop ends
    'c : current root

    ' Written by: Ryan Goh - 23 Oct 2023
    
    Dim c As Double
    Dim fa As Double
    Dim fc As Double
    Dim i As Long: i = 0
    Dim xold As Double
    
    'xr stores the list of c roots
    Dim xr As Object: Set xr = CreateObject("System.Collections.ArrayList")
    'Approximate error
    Dim ea As Object: Set ea = CreateObject("System.Collections.ArrayList")
    
    Do While (xu - xl) / 2 > tol
        c = (xl + xu) / 2
        If i = 0 Then
            fa = Application.Run(ThisWorkbook.Name & "!" & f, xl)
        End If
        fc = Application.Run(ThisWorkbook.Name & "!" & f, c)
        If fc = 0 Then
            Exit Do
        ElseIf fa * fc < 0 Then
            xu = c
        Else
            xl = c
            fa = fc
        End If
        ea.Add Abs((c - xold) / c)
        i = i + 1
        xold = c
        xr.Add c
    Loop
    
    If displayFullRes Then
        BisectionMethod = c
    Else
        BisectionMethod = CombineVector(1, ea.toarray, xr.toarray)
    End If
End Function

'Sub Main()
'    x = BisectionMethod("func", 1.5, 3, 0.005)
'    Debug.Print x
'End Sub

