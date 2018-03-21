Public Function ArchUnits(FeetInches As String) As Variant
    On Error GoTo ErrorHandler
    '**************************************************************************************************
    'This function will translate strings such as 10'-6" into
    'Architectural units as data type double in Feet
    'as 10.5000000000000
    '
    'If there is no ' or " in the string, the standard is thought
    'of as Inches.
    '
    'Accepted entries are:
    '   10'6"; 10'6; 10'-6"; 10'-6; 10'-6 1/2"; 10'-6 1/2; 10'; 6; 6-1/2; 6-1/2"; 6 1/2; 6 1/2"; 1/2;
    '   1/2"; 10.5'; 6.5"; 6.5
    '**************************************************************************************************
    
    Dim Feet As Double, Inches As Double
    Dim varData As Variant, varFraction As Variant
    Dim blnFeet As Boolean
    
    'Incase there are too many spaces
    FeetInches = Application.WorksheetFunction.Trim(FeetInches)
    
    'Check for any letters in the string
    If FeetInches Like "*[A-Z]*" Then
        ArchUnits = CVErr(xlErrValue)
        Exit Function
    End If
    
    'Check for more than 1 "-" or "/" or " "
    If (Len(FeetInches) - Len(Replace(FeetInches, "-", ""))) > 1 Or _
       (Len(FeetInches) - Len(Replace(FeetInches, "/", ""))) > 1 Or _
       (Len(FeetInches) - Len(Replace(FeetInches, " ", ""))) > 1 Then
       ArchUnits = CVErr(xlErrValue)
       Exit Function
    End If
    
    'Start off with examples of 10'-6" or 10'6" or 10'6
    If InStr(1, FeetInches, "'") <> 0 Then
        blnFeet = True
        'Split FeetInches text based on examples 10'-6" or 10'6"
        If InStr(1, FeetInches, "-") <> 0 Then
            varData = Split(FeetInches, "-")
        Else
            varData = Split(FeetInches, "'")
        End If
        
        'Get rid of the apostrophe, if any
        If InStr(1, varData(0), "'") <> 0 Then varData(0) = Replace(varData(0), "'", "")
        
        'Incase of the example '-6" (no number before the apostrophe)
        'if this IF statement is false, Feet = 0
        If varData(0) <> "" Then Feet = varData(0)
        
        'Check if the example was like 10' or 10'-
        If varData(1) = vbNullString Then
            ArchUnits = Feet
            Exit Function
        End If
        
        varData = varData(1)
    End If
    
    'If example 6-1/2 or 6-1/2"
    'and no feet portion of the data, only inches
    If blnFeet = False Then
        If InStr(1, FeetInches, "-") <> 0 Then
            varData = Split(FeetInches, "-")
            varData = varData(0) & " " & varData(1)
        Else
            If FeetInches = vbNullString Then
                ArchUnits = CVErr(xlErrValue)
                Exit Function
            End If
            
            varData = FeetInches
        End If
    End If
    
    'Check for " in the Inches portion of the string (varData(1))
    If InStr(1, varData, """") <> 0 Then varData = Replace(varData, """", vbNullString)
    
    'Check for spaces, then divisors for examples of 6 1/2
    If InStr(1, varData, " ") <> 0 Then
        'Check if there's a divisor. If not, exit function for example of 10'-6 1
        If InStr(1, varData, "/") = 0 Then ArchUnits = CVErr(xlErrValue): Exit Function
        
        varData = Split(varData, " ")
        varFraction = Split(varData(1), "/")
        Inches = varData(0) + (varFraction(0) / varFraction(1))
    Else
        'This is for examples of 10'6
        Inches = varData
    End If
    
            
    'JUST incase inches equals zero
    If Inches <> 0 Then ArchUnits = Feet + (Inches / 12) Else ArchUnits = Feet
    Exit Function
    
ErrorHandler:
    ArchUnits = CVErr(xlErrValue)
End Function