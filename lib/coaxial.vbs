Class TxLineCoaxial

  ''Define public variables
  Public length
  Public innerRadius
  Public outerRadius
  Public offsetX
  Public offsetY
  Public offsetZ
  Public orientation
  Public componentName
  Public solidName
  Public material

  ''Define private variables
  Private innerName
 

  ''Define public methods

  Public Sub Origin(X, Y, Z)
    offsetX = X
    offsetY = Y
    offsetZ = Z
  End Sub

  Public Sub Create
    OuterCylinder
    InnerCylinder

    With project.Solid 
      .Subtract componentName + ":" + solidName, componentName + ":" + innerName
    End With 

    TransformSolid componentName, solidName, orientation, offsetX, offsetY, offsetZ
  End Sub

  Public Function Impedance()
    charImpedance = 60*log((outerRadius)/(innerRadius))
    impedance = charImpedance
  End Function


  ''Define private methods

  Private Sub InnerCylinder()
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
  End Sub 

  Private Sub OuterCylinder()
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
  End Sub

End Class
