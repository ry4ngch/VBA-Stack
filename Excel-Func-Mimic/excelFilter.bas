Attribute VB_Name = "excelFilter"
Function xlFilter(ByVal rng As Range, ByVal searchVal As Range) As Variant
    ' This function imitate Office 365 filter() function using the search and filterXML function which is available in the older version of excel
    ' The function does a partial match as well
    ' Written by: Ryan Goh Chuang Hong - 29 Sept 2023
    If searchVal.Columns.Count > 1 Or searchVal.Rows.Count > 1 Then
        xlFilter = CVErr(xlErrValue)
        Exit Function
    End If
    
    If rng.Columns.Count > 1 And rng.Rows.Count = 1 Then
        filtered = Evaluate("FILTERXML(""<x><y>""&SUBSTITUTE(TEXTJOIN("","", TRUE, IFERROR(INDEX(" & rng.Address & ", N(IF(1, AGGREGATE(14,7,MATCH(" & rng.Address & ", IF(ISNUMBER(SEARCH(" & searchVal.Address & "," & rng.Cells.Address & "))," & rng.Address & "),0),COLUMN(" & rng.Address & ")-COLUMN(" & rng.Cells(1, 1).Address & ")+1)))),"""")),"","",""</y><y>"")&""</y></x>"",""//y"")")
        xlFilter = Application.Transpose(filtered)
    Else
        xlFilter = Evaluate("FILTERXML(""<x><y>""&SUBSTITUTE(TEXTJOIN("","", TRUE, IFERROR(INDEX(" & rng.Address & ", N(IF(1, AGGREGATE(14,7,MATCH(" & rng.Address & ", IF(ISNUMBER(SEARCH(" & searchVal.Address & "," & rng.Cells.Address & "))," & rng.Address & "),0),ROW(" & rng.Address & ")-ROW(" & rng.Cells(1, 1).Address & ")+1)))),"""")),"","",""</y><y>"")&""</y></x>"",""//y"")")
    End If
End Function

