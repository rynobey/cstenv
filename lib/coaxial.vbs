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
    .Axis orientation
    .Zrange cStr(offsetZ), cStr(offsetZ + length) 
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
    .Axis orientation
    .Zrange cStr(offsetZ), cStr(offsetZ + length) 
    .Xcenter "0" 
    .Ycenter "0" 
    .Segments "0" 
    .Create 
  End With 

  With project.Solid 
    .Subtract componentName + ":" + solidName, componentName + ":" + innerName
  End With 

End function
