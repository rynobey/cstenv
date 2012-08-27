Set obj = New CircularSectionLink

Class CircularSectionLink

  Public material
  Public componentName
  Public solidName

  Private CSTProject
  Private Solid

  Private Sub Class_Initialize
    componentName = "Coaxial"
    material = "Vacuum"
  End Sub

  Public Sub Init(NewCSTProject)
    Set CSTProject = NewCSTProject
  End Sub

  Public Sub CreateFromPick(circleRadius, angleX, angleY, angleZ, originX, originY, originZ)

    With CSTProject.Pick
      if angleX = 0 And angleZ = 0 then 
        .AddEdge cStr(originX), cStr(originY-10), cStr(originZ), cStr(originX), cStr(originY+10), cStr(originZ)
      End if
    End With

    'rotate face
    With CSTProject.Rotate 
      .Reset 
      .Component componentName
      solidName = CSTProject.Solid.getNextFreeName()
      .Name solidName 
      .Material material
      .Mode "Picks" 
      if angleX = 0 And angleZ = 0 then
        .Angle cStr(angleY)
      End if
      .Height "0.0" 
      .RadiusRatio "1.0" 
      .NSteps "0" 
      .SplitClosedEdges "True" 
      .SegmentedProfile "False" 
      .DeleteBaseFaceSolid "False" 
      .ClearPickedFace "True" 
      .SimplifySolid "True" 
      .UseAdvancedSegmentedRotation "True" 
      .Create 
    End With

  End Sub

End Class
