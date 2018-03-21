Option Explicit

Public CopyNewDataRange As Range

Public Sub CopyNewData()
    Dim copyTo As Range
    Set copyTo = Sheets("Sheet2").Range(CopyNewDataRange.Address)
    
    CopyNewDataRange.Copy copyTo
End Sub
