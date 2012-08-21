Set obj = New Script

Class Script
  
  Private fs
  Private text
  Private file

  Private Sub Class_Initialize
    Set fs = CreateObject("Scripting.FileSystemObject")
    text = ""
  End Sub

  Private Sub Class_Terminate
    On Error Resume Next
    file.Close
    Set file = Nothing
    Err.Clear
  End Sub

  Public Sub AddLine(input)
    text = text & vbNewLine & input
  End Sub

  Public Sub SaveAs(filePath)
    Set file = fs.OpenTextFile(filePath, 2, True, 0)
    file.write(text)
    file.Close
    Set file = Nothing
  End Sub

End Class
