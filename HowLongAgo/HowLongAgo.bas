Option Explicit

' Summary: Returns a friendly timestamp like "Just now" or "45 minutes ago"
' Param cell: The cell used for the function.
'             Should only be a single cell and should be a date with or without time.
Public Function HowLongAgo(ByRef cell As Range) As String
    ' Only allow dates and only allow single cells
    If ((Not IsDate(cell)) Or cell.Cells.Count > 1) Then
        HowLongAgo = vbNullString
        Exit Function
    End If
    
    Dim amount As Long
    
    amount = DateDiff("yyyy", cell, Now)

    If (amount > 0) Then
        HowLongAgo = GetHowLongAgoText(amount, "year")
        Exit Function
    End If

    amount = DateDiff("m", cell, Now)

    If (amount > 0) Then
        HowLongAgo = GetHowLongAgoText(amount, "month")
        Exit Function
    End If
    
    amount = DateDiff("d", cell, Now)
    
    If (amount > 0) Then
        HowLongAgo = GetHowLongAgoText(amount, "day")
        Exit Function
    End If
    
    If (Not HasTimeValue(cell.value)) Then
        HowLongAgo = "Today"
        Exit Function
    End If
    
    amount = DateDiff("h", cell, Now)
    
    If (amount > 0) Then
        HowLongAgo = GetHowLongAgoText(amount, "hour")
        Exit Function
    End If
    
    amount = DateDiff("n", cell, Now)
    
    If (amount > 0) Then
        HowLongAgo = GetHowLongAgoText(amount, "minute")
        Exit Function
    End If
    
    amount = DateDiff("s", cell, Now)
    
    If (amount >= 0) Then
        HowLongAgo = GetHowLongAgoText(amount, "second")
    Else
        HowLongAgo = "Hasn't happened yet..."
    End If
End Function

Private Function GetHowLongAgoText(ByVal amount As Long, ByRef unit As String) As String
    Select Case amount
        Case 0
            GetHowLongAgoText = "Just now"
        Case 1
            GetHowLongAgoText = "1 " & unit & " ago"
        Case 2
            GetHowLongAgoText = "A couple " & unit & "s ago"
        Case 3
            GetHowLongAgoText = "A few " & unit & "s ago"
        Case Else
            GetHowLongAgoText = amount & " " & unit & "s ago"
    End Select
End Function

Private Function HasTimeValue(ByRef value As String) As Boolean
    HasTimeValue = InStr(1, value, ":")
End Function
