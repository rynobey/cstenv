function txLineConical(radius, theta1, theta2, offset, orientation, componentName, solidName, material)

  ''Example input arguments
  'radius = 5
  'theta1 = 75
  'theta2 = 90
  'offset = 1
  'material = "PEC"
  'orientation = ""
  'componentName = "conicalTx"
  'solidName = "conicalLine"


  ''Internal settings
  dim profileName, endName
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

  With project.Transform 
    .Reset 
    .component componentName
    .Name componentName + ":" + profileName
    .Vector "0", "0", cStr(offset) 
    .UsePickedPoints "False" 
    .InvertPickedPoints "False" 
    .MultipleObjects "False" 
    .GroupObjects "False" 
    .Repetitions "1" 
    .MultipleSelection "False" 
    .Transform "Shape", "Translate" 
  End With 

  With project.Transform 
    .Reset 
    .component componentName
    .Name componentName + ":" + profileName
    .Origin "Free" 
    .Center "0", "0", cStr(offset) 
    .Angle "0", cStr(90-theta2+dTheta/2), "0"  
    .MultipleObjects "False" 
    .GroupObjects "False" 
    .Repetitions "1" 
    .MultipleSelection "False" 
    .Transform "Shape", "Rotate" 
  End With

  With project.Pick
    .ClearAllPicks
    .PickFaceFromId componentName + ":" + profileName, "1"
    .AddEdge "0.0", "0.0", cStr(offset + 10), "0.0", "0.0", cStr(offset - 10)
  End With

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

  project.Pick.ClearAllPicks
  project.Solid.Delete componentName + ":" + profileName

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
    .Center "0", "0", cStr(offset)
    .Segments "0" 
    .Create 
  End With

  project.Solid.Intersect componentName + ":" + solidName, componentName + ":" + endName 

end function
