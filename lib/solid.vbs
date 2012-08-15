Sub TransformSolid(componentName, solidName, orientation, offsetX, offsetY, offsetZ)
  ''Rotate to orientation
  With project.Transform 
    .Reset 
    .component componentName
    .Name componentName + ":" + solidName 
    .Origin "Free" 
    .Center "0", "0", "0" 
    if orientation = "x" then
      .Angle "0", "90", "0"
    elseif orientation = "y" then
      .Angle "-90", "0", "0"
    elseif orientation = "z" then
      .Angle "0", "0", "0"
    end if
    .MultipleObjects "False" 
    .GroupObjects "False" 
    .Repetitions "1" 
    .MultipleSelection "False" 
    .Transform "Shape", "Rotate" 
  End With 

  ''Move to offset
  With project.Transform 
    .Reset 
    .component componentName
    .Name componentName + ":" + solidName
    .Vector cStr(offsetX), cStr(offsetY), cStr(offsetZ) 
    .UsePickedPoints "False" 
    .InvertPickedPoints "False" 
    .MultipleObjects "False" 
    .GroupObjects "False" 
    .Repetitions "1" 
    .MultipleSelection "False" 
    .Transform "Shape", "Translate" 
  End With 
End Sub
