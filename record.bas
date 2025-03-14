Attribute VB_Name = "record"
Sub record()
    Dim logStr As String
    logStr = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    ActiveCell.offset(0, 0).Value = logStr & ": update "
End Sub
