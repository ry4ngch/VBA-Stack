Attribute VB_Name = "ModulusOperator"

Function XMod(ByVal Number As Double, ByVal Divisor As Double) As Double
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' XMod
    ' Performs the same function as Mod but will not overflow
    ' with very large numbers. Both Mod and integer division ( \ )
    ' will overflow with very large numbers. XMod will not.
    ' Existing code like:
    '       Result = Number Mod Divisor
    ' should be changed to:
    '       Result = XMod(Number, Divisor)
    ' Input values that are not integers are truncated to integers. Negative
    ' numbers are converted to postive numbers.
    ' This can be used in VBA code and can be called directly from
    ' a worksheet cell.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Number = Int(Abs(Number))
    Divisor = Int(Abs(Divisor))
    XMod = Number - (Int(Number / Divisor) * Divisor)
End Function







