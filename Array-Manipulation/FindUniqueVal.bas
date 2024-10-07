Attribute VB_Name = "FindUniqueVal"
Function ArrUnique(ByVal arr As Variant) As Variant
    ' This function accepts both range and arrays types and return the unique values
    ' It works only on 1-D rows or columns, vertical or horizontal data can be used.
    ' Written by: Ryan Goh Chuang Hong - 5 Oct 2023

    Dim isTwoD As Boolean: isTwoD = getDims(arr) > 1
    If TypeName(arr) = "Range" Then
        Dim orient As Boolean: orient = arr.Rows.Count > 1
        arr = arr.Value2
    End If
    
    Dim item As Variant
    
    Dim uniqueList As Object
    #If Mac Then
        Set uniqueList = New ArrayList
    #Else
        Set uniqueList = CreateObject("System.Collections.ArrayList")
    #End If
    
    For Each item In arr
        If Not IsError(item) Then
            If Not IsInArray(item, uniqueList.ToArray) Then
                uniqueList.Add item
            End If
        End If
    Next item
    
    If orient Or isTwoD Then
        ArrUnique = Application.Transpose(uniqueList.ToArray)
    Else
        ArrUnique = uniqueList.ToArray
    End If
End Function
