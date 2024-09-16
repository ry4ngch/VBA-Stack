Attribute VB_Name = "ArrDimension"
Option Explicit

Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Integer)

Public Function GetDims(VarSafeArray As Variant) As Integer
    ' This function returns the dimensions of an array
    ' Written by: Blackhawk - 24 Oct 2014
    ' https://stackoverflow.com/questions/24613101/vba-check-if-array-is-one-dimensional/26555865#26555865
    
    ' The function has been tested to work on VBA arrays. It however does not work with range and cannot be used directly in excel
    ' Modification has been made by checking for input type if range then converts to array
    ' From testing, it seems that if the 1-D range is horizontal, it will be interpreted as 1-D array while if the 1-D range is vertical, it will be interpreted as 2-D array.
    ' Modified by: Ryan Goh - 26 Sept 2023
    
    ' Output results: 1 for 1-D Array, 2 for 2-D Array. The function returns 0 if mismatch or if no array is initialize.
    
    
    If TypeName(VarSafeArray) = "Range" Then
        VarSafeArray = Application.Transpose(Application.Transpose(VarSafeArray.Value2))
    End If


    Dim variantType As Integer
    Dim pointer As LongPtr
    Dim arrayDims As Integer

    CopyMemory VarPtr(variantType), VarPtr(VarSafeArray), 2& 'the first 2 bytes of the VARIANT structure contain the type

    If (variantType And &H2000) > 0 Then 'Array (&H2000)
        'If the Variant contains an array or ByRef array, a pointer for the SAFEARRAY or array ByRef variant is located at VarPtr(VarSafeArray) + 8
        CopyMemory VarPtr(pointer), VarPtr(VarSafeArray) + 8, LenB(pointer)

        'If the array is ByRef, there is an additional layer of indirection through another Variant (this is what allows ByRef calls to modify the calling scope).
        'Thus it must be dereferenced to get the SAFEARRAY structure
        If (variantType And &H4000) > 0 Then 'ByRef (&H4000)
            'dereference the pointer to pointer to get the actual pointer to the SAFEARRAY
            CopyMemory VarPtr(pointer), pointer, LenB(pointer)
        End If
        'The pointer will be 0 if the array hasn't been initialized
        If Not pointer = 0 Then
            'If it HAS been initialized, we can pull the number of dimensions directly from the pointer, since it's the first member in the SAFEARRAY struct
            CopyMemory VarPtr(arrayDims), pointer, 2&
            GetDims = arrayDims
        Else
            GetDims = 0 'Array not initialized
        End If
    Else
        GetDims = 0 'It's not an array... Type mismatch maybe?
    End If
End Function