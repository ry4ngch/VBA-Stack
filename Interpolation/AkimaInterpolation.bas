Attribute VB_Name = "AkimaInterpolation"
Option Explicit
Option Base 0

Public Function Akima(ByVal known_y As Variant, ByVal known_x As Variant, ByVal interp_values As Variant) As Variant
    '+-----------------------------------------------------------------------------------------+'
    '| Adapted from: https://asmquantmacro.com/2015/09/01/akima-spline-interpolation-in-excel/ |'
    '+-----------------------------------------------------------------------------------------+'
    
    '+--------------------------------------------------------------------------------------------+
    '| check if selected range is vertical or horizontal in order to return the right orientation |
    '+--------------------------------------------------------------------------------------------+
    Dim isRangeVert As Boolean
    Dim interpType As String: interpType = TypeName(interp_values)
    If TypeName(interp_values) = "Range" Then
        ' if True then is vertical else it is horizontal
        isRangeVert = (interp_values.Columns.Count = 1)
    End If
    
    '+------------------------------------------------------------------------------+'
    '| Convert the array from range to base 0 since range arrays start from index 1 |'
    '+------------------------------------------------------------------------------+'
    If TypeName(known_x) = "Range" Then
        known_x = CVArrayBase(Application.Transpose(known_x.Value2))
        known_y = CVArrayBase(Application.Transpose(known_y.Value2))
    End If
    Dim n As Integer: n = UBound(known_x) - LBound(known_x) + 1
    
    Dim ii As Integer: ii = 0
    ReDim m(n + 3) As Double
    '+--------------------------------------------------------+'
    '| shift data by + 2 in the array and compute the secants |'
    '| also calculate extrapolated end point secants          |'
    '+--------------------------------------------------------+'

    For ii = LBound(known_x) To UBound(known_x) - 1
        m(ii + 2) = (known_y(ii + 1) - known_y(ii)) / (known_x(ii + 1) - known_x(ii))
    Next ii

    m(1) = 2 * m(2) - m(3)
    m(0) = 2 * m(1) - m(2)
    m(n + 1) = 2 * m(n) - m(n - 1)
    m(n + 2) = 2 * m(n + 1) - m(n)
    '+--------------------------------------------------------+'
    '| Calculate slope at each data point                     |'
    '+--------------------------------------------------------+'

    Dim A As Double
    Dim B As Double
    ReDim t(UBound(known_x)) As Double

    For ii = LBound(known_x) To UBound(known_x)
        A = Abs(m(ii + 3) - m(ii + 2))
        B = Abs(m(ii + 1) - m(ii))
        If (A + B) <> 0 Then
            t(ii) = (A * m(ii + 1) + B * m(ii + 2)) / (A + B)
        Else
            t(ii) = 0.5 * (m(ii + 2) + m(ii + 1))
        End If
    Next ii

    '+----------------------------------------------------------------+'
    '| For each value we wish to interpolate locate the spline segment|'
    '| and calculate the coefficients                                 |'
    '+----------------------------------------------------------------|'

    Dim intTop As Integer
    Dim intBottom As Integer
    Dim intMiddle As Integer
    
    If TypeName(interp_values) = "Range" Then
        If interp_values.Count > 1 Then
            interp_values = CVArrayBase(Application.Transpose(interp_values.Value2))
        Else
            interp_values = Array(interp_values)
        End If
    ElseIf TypeName(interp_values) = "Double" Or TypeName(interp_values) = "Long" Or TypeName(interp_values) = "Integer" Then
        interp_values = Array(interp_values)
    End If
    
    
    ReDim output(UBound(interp_values)) As Variant
    For ii = LBound(interp_values) To UBound(interp_values)
        '+----------------------------------------------------------------+'
        '| Binary (bisection) search for the interpolating interval for x |'
        '+----------------------------------------------------------------+'

        intBottom = LBound(known_x)
        intTop = UBound(known_x)
        While (intTop - intBottom) > 1
            intMiddle = Fix(0.5 * (intBottom + intTop))
            If known_x(intMiddle) > interp_values(ii) Then
                intTop = intMiddle
            Else
                intBottom = intMiddle
            End If
        Wend

        B = known_x(intTop) - known_x(intBottom)
        If B = 0 Then
            Akima = "Bad x input"
        End If

        A = interp_values(ii) - known_x(intBottom)
        output(ii) = known_y(intBottom) + t(intBottom) * A + (3 * m(intBottom + 2) - _
                        2 * t(intBottom) - t(intBottom + 1)) * A * A / B + (t(intBottom) + t(intBottom + 1) - _
                        2 * m(intBottom + 2)) * A * A * A / (B * B)

    Next ii
    
    If interpType = "Range" Then
        If isRangeVert Then
            Akima = Application.Transpose(output)
        Else
            Akima = output
        End If
    Else
        Akima = output
    End If
    
End Function


