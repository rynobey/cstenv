Class TxLineConical

  ''Define public variables
  Public radius
  Public theta1
  Public theta2
  Public offsetX
  Public offsetY
  Public offsetZ
  Public orientation
  Public componentName
  Public solidName
  Public material

  ''Define private variables
  Private profileName, endName


  ''Public methods 

  Public Sub Origin(X, Y, Z)
    offsetX = X
    offsetY = Y
    offsetZ = Z
  End Sub

  Public Function Create()
    CreateFace
    TransformFace
    RotateFace

    ''Remove temp objects
    project.Pick.ClearAllPicks
    project.Solid.Delete componentName + ":" + profileName

    SphericalEnd
    TransformSolid componentName, solidName, orientation, offsetX, offsetY, offsetZ
  End Function 


  ''Private methods 

  Private Sub CreateFace()
    ''Internal settings
    dTheta = abs(theta1 - theta2)
    With project.AnalyticalFace
      .Reset 
      .Component componentName 
      profileName = project.Solid.getNextFreeName()
      .Name profileName
      .Material material
      .LawX "u"
      .LawY "0"
      .LawZ "v*u" + "*" + cStr(project.tanD(dTheta/2))
      .ParameterRangeU "0", radius
      .ParameterRangeV "-1", "1" 
      .Create
    End With
  End Sub

  Private Sub TransformFace()
    ''Internal settings
    dTheta = abs(theta1 - theta2)
    With project.Transform 
      .Reset 
      .component componentName
      .Name componentName + ":" + profileName
      .Origin "Free" 
      .Center "0", "0", "0" 
      .Angle "0", cStr(90-theta2+dTheta/2), "0"  
      .MultipleObjects "False" 
      .GroupObjects "False" 
      .Repetitions "1" 
      .MultipleSelection "False" 
      .Transform "Shape", "Rotate" 
    End With
  End Sub

  Private Sub RotateFace()
    ''Pick Face and add edge for rotation
    With project.Pick
      .ClearAllPicks
      .PickFaceFromId componentName + ":" + profileName, "1"
      .AddEdge "0.0", "0.0", "10", "0.0", "0.0", "-10"
    End With

    ''Rotate face to create conical transmission line
    With project.Rotate 
      .Reset 
      .Component componentName
      .Name solidName
      .Material material
      .Mode "Picks" 
      .Angle "360" 
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

  Private Sub SphericalEnd()
    ''Intersect with a sphere to create spherical ends
    With project.Sphere 
      .Reset 
      .Component componentName
      endName = project.Solid.getNextFreeName()
      .Name endName
      .Material material 
      .Axis "z" 
      .CenterRadius radius
      .TopRadius "0" 
      .BottomRadius "0" 
      .Center "0", "0", "0"
      .Segments "0" 
      .Create 
    End With
    arg1 = componentName + ":" + solidName
    arg2 = componentName + ":" + endName
    project.Solid.Intersect arg1, arg2
  End Sub

End Class
