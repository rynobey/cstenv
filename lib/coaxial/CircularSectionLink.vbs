Set obj = New CircularSectionLink

Class CircularSectionLink

  Private CSTProject
  Private Solid

  Public Sub Init(NewCSTProject)
    Set CSTProject = NewCSTProject
  End Sub

  Public Sub CreateFromPick(circleRadius, angleX, angleY, angleZ, originX, originY, originZ)

    With project.Pick
      if angleX = 0 And angleZ = 0 then 
        .AddEdge cStr(originX), cStr(originY-10), cStr(originZ), cStr(originX), cStr(originY+10), cStr(originZ)
      End if
    End With

    'rotate face
    With project.Rotate 
      .Reset 
      .Component componentName
      primerSolidName = project.Solid.getNextFreeName()
      .Name primerSolidName 
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
