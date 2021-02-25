# vba-excel
//when I enter a percentage in cell. H8 it give the value in cell. L8, but also I want it to be able to do the opposite. When I enter a value in cell. L8 it give the percentage in cell. H8. Meaning that the user will have a choice to enter either the percentage (H8) or the value (L8).
pre
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("H8")) Is Nothing Then
      If Range("H8").HasFormula Then Exit Sub
      Range("L8").Value = "=H8*L7/100"
    End If
    If Not Intersect(Target, Range("L8")) Is Nothing Then
      If Range("L8").HasFormula Then Exit Sub
      Range("H8").Value = "=L8/L7*100"
    End If
End Sub
/pre
