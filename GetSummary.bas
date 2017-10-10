Option Explicit

Public Function GetSummary(text As String, num_of_words As Long) As String
    If (num_of_words <= 0) Then
        GetSummary = ""
        Exit Function
    End If

    Dim words() As String
    words = Split(text, " ")
    
    Dim result As String
    result = words(0)

    Dim i As Long
    i = 1
    Do While (i < num_of_words)
        result = result & " " & words(i)
        i = i + 1
    Loop

    GetSummary = result & "..."
End Function
