Attribute VB_Name = "ConvertArrayBase"

Function CVArrayBase(ByVal list As Variant) As Variant
    ''' This function will convert 1D array to base 0, so first element index is 0
    ''' Written by: Ryan Goh - 3 Oct 2023
    
    Dim tempArr As Object: Set tempArr = CreateObject("System.Collections.ArrayList")
    Dim d As Variant
    
    For Each d In list
        tempArr.Add (d)
    Next d
    
    CVArrayBase = tempArr.toarray()
End Function

