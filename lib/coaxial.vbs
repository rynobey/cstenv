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

  Public Function AddPort(number)
    With project.Port 
      .Reset 
      .PortNumber cStr(number)
      .Label "" 
      .NumberOfModes "1" 
      .AdjustPolarization "False" 
      .PolarizationAngle "0.0" 
      .ReferencePlaneDistance "0" 
      .TextSize "50" 
      .Coordinates "Free" 
      .Orientation "zmax" 
      .PortOnBound "False" 
      .ClipPickedPortToBound "False" 
      .Xrange cStr(-OuterRadius), cStr(OuterRadius) 
      .Yrange cStr(-OuterRadius), cStr(OuterRadius) 
      .Zrange cStr(Length/2), cStr(Length/2) 
      .XrangeAdd "0.0", "0.0" 
      .YrangeAdd "0.0", "0.0" 
      .ZrangeAdd "0.0", "0.0" 
      .SingleEnded "False" 
      .Create 
    End With 
 
    With project.Transform 
      .Reset 
      .Name "port" + cStr(number)
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
      .Transform "Port", "Rotate" 
    End With 

    With project.Transform 
      .Reset 
      .Name "port" + cStr(number) 
      .Vector cStr(offsetX), cStr(offsetY), cStr(offsetZ)
      .UsePickedPoints "False" 
      .InvertPickedPoints "False" 
      .MultipleObjects "False" 
      .GroupObjects "False" 
      .Repetitions "1" 
      .MultipleSelection "False" 
      .Transform "Port", "Translate" 
    End With 

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
