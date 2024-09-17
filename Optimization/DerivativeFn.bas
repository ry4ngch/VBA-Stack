Attribute VB_Name = "DerivativeFn"
Function SolveDerivative(ByVal coefficients As Variant, ByVal powers As Variant, Optional ByVal variable As String) As String
    ' Inputs:
    '   - coefficients: key in the coefficients of the equation as a range or array
    '   - powers: key in the powers of each term of the equation as a range or array
    '   - variable: input the string variable representing the symbol of the equation
    
    ' Written by: Ryan Goh - 17 Sept 2024
    
    Dim i As Long
    Dim coeff As Double
    Dim power As Double
    Dim term As String
    Dim result As String
    
    ' Check if variable argument is input, else assign "x" as default
    If StrPtr(variable) = 0 Then
        variable = "x"
    End If
    
    ' *** Check if arguments are of type "Range", means that the inputs are from Excel Worksheet. ***
    If TypeName(coefficients) = "Range" Then
        ' If coefficients is column wise, we only need to tranpose once
        If coefficients.Columns.Count = 1 Then
            coefficients = Application.Transpose(coefficients.Value2)
        Else ' If coefficients is row wise, we only need to tranpose twice
            coefficients = Application.Transpose(Application.Transpose(coefficients.Value2))
        End If
    End If
    
    If TypeName(powers) = "Range" Then
    '    If powers is column wise, we only need to tranpose once
        If powers.Columns.Count = 1 Then
            powers = Application.Transpose(powers.Value2)
        Else '' If coefficients is row wise, we only need to tranpose twice
            powers = Application.Transpose(Application.Transpose(powers.Value2))
        End If
    End If
    
    ' *** Compute the derivative ***
    result = ""

    ' Loop through the arrays of coefficients and powers
    For i = LBound(coefficients) To UBound(coefficients)
        coeff = coefficients(i)
        power = powers(i)

        ' Skip term if power is 0 since its derivative is 0
        If power <> 0 Then
            ' Derivative: new coefficient = old coefficient * power
            coeff = coeff * power

            ' New power = old power - 1
            power = power - 1

            ' Create term string based on new coefficient and power
            If power = 0 Then
                term = CStr(coeff)  ' If power is 0, it's just the constant
            ElseIf power = 1 Then
                term = CStr(coeff) & variable  ' If power is 1, just x
            Else
                term = CStr(coeff) & variable & "^" & CStr(power)
            End If

            ' Concatenate terms
            If result = "" Then
                result = term
            Else
                If coeff > 0 Then
                    result = result & " + " & term
                Else
                    result = result & " " & term
                End If
            End If
        End If
    Next i

    ' Return the derivative equation as a string
    SolveDerivative = result
End Function


Sub TestDerivative()
    Dim coefficients As Variant
    Dim powers As Variant
    Dim result As String
    
    ' Example: equation 3x^4 + 5x^3 + 2x
    ' coefficients = Array(3, 5, 2)
    ' powers = Array(4, 3, 1)
    
    ' Example2: equation 4y^5 + 3y^2 - 2y
'    coefficients = Array(4, 3, -2)
'    powers = Array(5, 2, 1)
'
    'Example 3:equation 8z^9 + 4z^5 +8z + 3
    coefficients = Array(8, 4, 8, 3)
    powers = Array(9, 5, 1, 0)
    
    result = SolveDerivative(coefficients, powers, "z")
    MsgBox "The derivative is: " & result
    
    
End Sub
