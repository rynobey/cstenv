Set obj = New CST 

Class CST
  
  Public CSTCOM

  Private Sub Class_Initialize
    Set CSTCOM = CreateObject("CSTStudio.Application")
  End Sub

  Private Sub Class_Terminate
    Set CSTCOM = Nothing
  End Sub
End Class
