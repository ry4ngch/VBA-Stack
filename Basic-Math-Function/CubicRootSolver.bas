Attribute VB_Name = "CubicRootSolver"
Function FindCubicRoot(Coef As Variant, Optional DisplayImgRoot As Boolean = True) As Variant
    ' Written by: Ryan Goh - 1/2/2023
    ' This function is written based on equations developed by Alex Tomanovich "POLYNOM - SOLVE Root of Eqn.xls"
    
    ' Arguments:
    '   - Coef: List the coefficients of the equations as Range or Array
    '   - DisplayImgRoot: Set this to True if we want the function to return the imaginary roots, by default this will be True.
    
    ' Declaration
    Dim A As Double
    Dim B As Double
    Dim c As Double
    Dim d As Double
    Dim f As Double
    Dim g As Double
    Dim H As Double
    Dim i As Double
    Dim j As Double
    Dim k As Double
    Dim L As Double
    Dim m As Double
    Dim n As Double
    Dim p As Double
    Dim r As Double
    Dim S As Double
    Dim t As Double
    Dim U As Double
    Dim x As Double
    Dim x1 As Double
    Dim x2 As Double
    Dim x3 As Double
    Dim OrientVert As Boolean
    Dim output As Variant
    
    If TypeName(Coef) = "Range" Then
        OrientVert = (Coef.Columns.Count = 1)
        Coef = CVArrayBase(Application.Transpose(Coef))
    End If
    
    If UBound(Coef) <> 3 Then
        FindCubicRoot = CVErr(xlErrNA)
        Exit Function
    End If
    
    A = Coef(0)
    B = Coef(1)
    c = Coef(2)
    d = Coef(3)
    
    f = ((3 * c / A) - (B ^ 2 / (A ^ 2))) / 3
    g = ((2 * B ^ 3 / (A ^ 3)) - (9 * B * c / (A ^ 2)) + (27 * d / A)) / 27
    H = (g ^ 2 / 4) + (f ^ 3 / 27)
    
    If H > 0 Then
        ' There is only 1 root real
        r = -(g / 2) + H ^ (1 / 2)
        S = r ^ (1 / 3)
        t = -(g / 2) - H ^ (1 / 2)
        U = Application.WorksheetFunction.Power(t, 1 / 3)
        x1 = (S + U) - (B / (3 * A))
        x2 = -(S + U) / 2 - (B / (3 * A))
        x2_img = (S - U) * Sqr(3) / 2
        x3 = -(S + U) / 2 - (B / (3 * A))
        x3_img = -(S - U) * Sqr(3) / 2
        
        If DisplayImgRoot Then
            output = Array(x1, WorksheetFunction.Complex(x2, x2_img), WorksheetFunction.Complex(x3, x3_img))
        Else
            output = Array(x1, x2, x3, x2_img, x3_img)
        End If
    Else
        If f = 0 And g = 0 And H = 0 Then
            ' All 3 roots are real and equal
            x = (d / A) ^ (1 / 3) * (-1)
            x1 = x
            x2 = x
            x3 = x
            output = Array(x1, x2, x3)
        Else
            ' All 3 roots are real
            i = Sqr((g ^ 2 / 4) - H)
            j = i ^ (1 / 3)
            k = Application.Acos(-(g / (2 * i)))
            L = j * -1
            m = Cos(k / 3)
            n = Sqr(3) * Sin(k / 3)
            p = (B / (3 * A)) * (-1)
            x1 = 2 * j * Cos(k / 3) - (B / (3 * A))
            x2 = L * (m + n) + p
            x3 = L * (m - n) + p
            output = Array(x1, x2, x3)
        End If
    
    End If
    
    If OrientVert Then
        FindCubicRoot = Application.Transpose(output)
    Else
        FindCubicRoot = output
    End If
    
    
End Function
