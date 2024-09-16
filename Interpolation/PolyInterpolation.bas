Attribute VB_Name = "PolyInterpolation"
Option Base 0

Static Function Log10(ByVal X As Double) As Double
    Log10 = Log(X) / Log(10#)
End Function

Function PolyInterpolate(ByVal xtab As Variant, ByVal ytab As Variant, ByVal X As Variant, Optional ByVal Order As Long, Optional ByVal extrapolateLim As Double = 0.2, Optional isSorted As Boolean = True) As Variant
    ' Input:
    '   xtab:  array of n x values
    '   ytab:  array of n y values
    '      x:  interpolant
    ' Order:
    '     n=0: Linear Logarithm
    '     n=1: linear
    '     n=2: quadratic
    '     n>2: average of two possible quadratics through nearest 4 local points
    
    ' Output:
    '      y:  interpolated y at x
    
    Dim Descending As Boolean
    Dim nx, lo, hi, mid
    Dim XVal As Variant
    Dim b1, b2 As Double
    Dim delhilo, dellolo1, delhilo1, delhi1hi, delhi1lo As Double
    Dim sortedData As Variant
    
    ' Default is more accurate quadratic interpolation
    ' Linear interpolation is for exceptional cases only
    If IsMissing(Order) Then Order = 2
  
    ' Convert range to array
    If TypeName(xtab) = "Range" Then
        xtab = CVArrayBase(Application.Transpose(xtab.Value2))
        ytab = CVArrayBase(Application.Transpose(ytab.Value2))
    End If
    
    If TypeName(X) = "Range" Then
        If X.Count > 1 Then
            XVal = CVArrayBase(Application.Transpose(X.Value2))
        Else
            XVal = Array(X)
        End If
    ElseIf TypeName(X) = "Double" Or TypeName(X) = "Long" Or TypeName(X) = "Integer" Then
        XVal = Array(X)
    Else
        XVal = X
    End If
    
    ReDim output(0 To UBound(XVal)) As Variant
    ReDim quad(0 To UBound(XVal)) As Variant
    ReDim nquad(0 To UBound(XVal)) As Variant

    ' Check Input Data
    ' Avoid loop counting data as it is inefficient for large tables
    nx = UBound(xtab)
    
    
    ' Error Handling
    ' # or we can return the number without interpolation here
    If (nx < 1) Then
        PolyInterpolate = "Error: nX<1"
        Exit Function
    End If
    
    If (nx <> UBound(ytab)) Then
      PolyInterpolate = "Error: nX<>nY"
      Exit Function
    End If
    
    ' Loop through all X Inputs to interpolate
    For i = LBound(XVal) To UBound(XVal)
        If Not isSorted Then
            sortedData = QuickSortedMatrix(CombineVector(1, xtab, ytab))
            xtab = sliceArray(sortedData, 0, 0, , , , 0)
            ytab = sliceArray(sortedData, 0, 1)
        End If
    
        ' Allow limited extrapolation
        If (((xtab(0) - XVal(i)) / (xtab(1) - xtab(0))) > extrapolateLim Or _
        ((XVal(i) - xtab(nx)) / (xtab(nx) - xtab(nx - 1))) > extrapolateLim) Then
            output(i) = "Error: >" & extrapolateLim * 100 & "% extrapolation"
            GoTo nextIteration
        End If

        ' Locate interval containing x with bisection
        ' Run time is very important for huge spreadsheet tables:
        '  a) Do not test for table order(via descending Xor...) inside Loop
        '  b) Do not check equality inside Loop
        ' Run time is not important at all for small tables:
        '  a) Accept 1 extra iteration to avoid equality check in loop
        Descending = (xtab(0) > xtab(nx))
        lo = 0
        hi = nx
        Do
            mid = Int((lo + hi) \ 2)
            If xtab(mid) < XVal(i) Then
                lo = mid
            Else
                hi = mid
            End If
        Loop Until Abs(hi - lo) <= 1
    
        ' Check for divide by zero and save result for re-use.
        delhilo = (xtab(hi) - xtab(lo))
        If (delhilo = 0) Then
            output(i) = "Error: Equal X values"
            GoTo nextIteration
        End If
        
        ' Linear for 2 points
        ' 3 values -> For eg: base 0 array, will have nx < 2
        If nx < LBound(xtab) + 2 Or Order < 2 Then
            If Order = 0 Then ' logarithm interpolation
                output(i) = 10 ^ (Log10(ytab(lo)) + (Log10(XVal(i)) - Log10(xtab(lo))) / (Log10(xtab(hi)) - Log10(xtab(lo))) * (Log10(ytab(hi)) - Log10(ytab(lo))))
            Else ' linear interpolation
                output(i) = ytab(lo) + (XVal(i) - xtab(lo)) / delhilo * (ytab(hi) - ytab(lo))
            End If
        Else
            ' when Order >= 2 only run newton quadractic polynomial
            ' Compute possible Newton quadratic polynomials
            '+-------------------------------------------------------------------+
            '| See: https://www.youtube.com/watch?v=2dWcFuJ09GQ&list=WL&index=15 |
            '| See: https://www.youtube.com/watch?v=cWbX8sLXDWo&list=WL&index=16 |
            '+-------------------------------------------------------------------+
            
            ' The second order polynomial is of the form:
            '   f2(x) = b0 + b1*(x(i) - x(0)) + b2*(x(i) - x(0))*(x(i) - x(1))
            '   whereby:
            '   b0 = f2(x(0))
            '   b1 = (f2(x(1)) - f2(x(0))) / (x(1) - x(0))
            '   b2 = ((f2(x(2)) - f2(x(1))) / (x(2) - x(1)) - (f2(x(1)) - f2(x(0))) / (x(1) - x(0)))/(x(2) - x(0))
            '   b2 = ((f2(x(2)) - f2(x(1))) / (x(2) - x(1)) - b1)/(x(2) - x(0))
            
            nquad(i) = 0
            quad(i) = 0
            If lo > LBound(xtab) Then
                dellolo1 = (xtab(lo) - xtab(lo - 1))
                If (dellolo1 = 0) Then
                  output(i) = "Error: Equal X values"
                  GoTo nextIteration
                End If
        
                delhilo1 = (xtab(hi) - xtab(lo - 1))
                If (delhilo1 = 0) Then
                  output(i) = "Error: Equal X values"
                  GoTo nextIteration
                End If
        
                b1 = (ytab(lo) - ytab(lo - 1)) / dellolo1
                b2 = ((ytab(hi) - ytab(lo)) / delhilo - b1) / delhilo1
                quad(i) = ytab(lo - 1) + b1 * (XVal(i) - xtab(lo - 1)) + b2 * (XVal(i) - xtab(lo - 1)) * (XVal(i) - xtab(lo))
                nquad(i) = nquad(i) + 1
            End If
            
           
            If hi < nx Then
                delhi1hi = (xtab(hi + 1) - xtab(hi))
                If (delhi1hi = 0) Then
                  output(i) = "Error: Equal X values"
                  If i = UBound(XVal) Then Exit For
                End If
        
                delhi1lo = (xtab(hi + 1) - xtab(lo))
                If (delhi1lo = 0) Then
                  output(i) = "Error: Equal X values"
                  If i = UBound(XVal) Then Exit For
                End If
        
                b1 = (ytab(hi) - ytab(lo)) / delhilo
                b2 = ((ytab(hi + 1) - ytab(hi)) / delhi1hi - b1) / delhi1lo
                quad(i) = quad(i) + ytab(lo) + b1 * (XVal(i) - xtab(lo)) + b2 * (XVal(i) - xtab(lo)) * (XVal(i) - xtab(hi))
                nquad(i) = nquad(i) + 1
            End If
            output(i) = quad(i) / nquad(i)
        End If
        
nextIteration:
        ' Do nothing here
    Next i
    
    If TypeName(X) = "Range" Then
        If X.Columns.Count = 1 Then
            PolyInterpolate = Application.Transpose(output)
        Else
            PolyInterpolate = output
        End If
    Else
        PolyInterpolate = output
    End If
    
End Function
