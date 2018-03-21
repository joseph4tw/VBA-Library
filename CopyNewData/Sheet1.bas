Option Explicit

Private Sub Worksheet_Change(ByVal target As Range)
    Dim sh As Worksheet
    Set sh = ActiveSheet
    
    Dim newdata As Range
    Set newdata = Application.Intersect(target, sh.Range("A2", "A" & sh.Rows.Count))

    If (Not newdata Is Nothing) Then
        If (CopyNewDataRange Is Nothing) Then
            Set CopyNewDataRange = newdata
        Else
            Set CopyNewDataRange = Application.Union(CopyNewDataRange, newdata)
        End If
    End If
End Sub
