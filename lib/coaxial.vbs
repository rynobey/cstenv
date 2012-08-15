function txLineCoaxial(length, innerRadius, outerRadius, offsetX, offsetY, offsetZ, orientation, componentName, solidName, material)

  ''Internal settings
  dim innerName
  
  With project.Cylinder 
    .Reset 
    innerName = project.Solid.getNextFreeName()
    .Name innerName 
    .Component componentName
    .Material material
    .OuterRadius innerRadius
    .InnerRadius "0"
    .Axis "z"
    .Zrange cStr(-length/2), cStr(length/2) 
    .Xcenter "0" 
    .Ycenter "0" 
    .Segments "0" 
    .Create 
  End With 

  With project.Cylinder 
    .Reset
    .Name solidName
    .Component componentName
    .Material material
    .OuterRadius outerRadius
    .InnerRadius "0"
    .Axis "z"
    .Zrange cStr(-length/2), cStr(length/2) 
    .Xcenter "0" 
    .Ycenter "0" 
    .Segments "0" 
    .Create 
  End With 

  With project.Solid 
    .Subtract componentName + ":" + solidName, componentName + ":" + innerName
  End With 

  ''rotate to orientation
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

End function
