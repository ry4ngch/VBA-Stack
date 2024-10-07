Attribute VB_Name = "ElementInArray"
Function IsInArray(ByVal VarToBeFound As Variant, ByVal arr As Variant) As Boolean
    ' This function is used to check if an element exist in an array
    Dim Element As Variant
    For Each Element In arr
        If Not (IsError(Element)) Then
            If Element = VarToBeFound Then
                IsInArray = True
                Exit Function
            End If
        End If
    Next Element

    IsInArray = False
End Function
