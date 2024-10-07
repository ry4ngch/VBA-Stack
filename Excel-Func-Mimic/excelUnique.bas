Attribute VB_Name = "excelUnique"
Function xlUnique(ByVal rng As Range) As Variant
    ' This function imitate Office 365 Unique() function
    ' Written by: Ryan Goh Chuang Hong - 2 Oct 2023
    
    If rng.Columns.Count > 1 And rng.Rows.Count = 1 Then
        uniqueVal = Evaluate("FILTERXML(""<x><y>""&SUBSTITUTE(TEXTJOIN("","",TRUE,IF(COLUMN(" & rng.Address & ")-COLUMN(" & rng.Cells(1, 1).Address & ")+1=MATCH(" & rng.Address & "," & rng.Address & ",0)," & rng.Address & ","""")),"","",""</y><y>"")&""</y></x>"",""//y"")")
        xlUnique = Application.Transpose(uniqueVal)
    Else
        xlUnique = Evaluate("FILTERXML(""<x><y>""&SUBSTITUTE(TEXTJOIN("","",TRUE,IF(ROW(" & rng.Address & ")-ROW(" & rng.Cells(1, 1).Address & ")+1=MATCH(" & rng.Address & "," & rng.Address & ",0)," & rng.Address & ","""")),"","",""</y><y>"")&""</y></x>"",""//y"")")
    End If
    
End Function

