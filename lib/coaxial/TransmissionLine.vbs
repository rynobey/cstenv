Set obj = New CoaxialTransmissionLine

Class CoaxialTransmissionLine 

  ''Define public variables
  Public length
  Public innerRadius
  Public outerRadius
  Public offsetX
  Public offsetY
  Public offsetZ
  Public angleX
  Public angleY
  Public angleZ
  Public orientation
  Public componentName
  Public solidName
  Public material
  Public charImpedance

  ''Define private variables
  Private innerName
  Private CSTProject
  Private Solid
 
  Private Sub Class_Initialize
    'set up default values
    charImpedance = 50
    OuterRadius = 2.3
    offSetX = 0
    offSetY = 0
    offSetZ = 0
    angleX = 0
    angleY = 0
    angleZ = 0
    orientation = "z"
    material = "Vacuum"
    componentName = "Coaxial"
    solidName = "Tx"
  End Sub


  ''Define public methods
  Public Sub Init(NewCSTProject)
    Set CSTProject = NewCSTProject
    Set Solid = Env.Use("lib\utils\Solid")
    Solid.Init(CSTProject)
  End Sub

  Public Function InnerFromOuter(impedance, radius)
    InnerFromOuter = radius/exp(impedance/60)
  End Function

  Public Function OuterFromInner(impedance, radius)
    OuterFromOuter = exp(impedance/60)*radius
  End Function

  Public Sub PickPositiveFace()
    'rotation in the x-direction:
    'if orientation = "z" then
    xr = length/2
    yr = xr*CSTProject.CosD(angleX)
    temp1 = xr*CSTProject.SinD(angleX)*xr*CSTProject.SinD(angleX)
    temp2 = yr*CSTProject.SinD(angleY)*yr*CSTProject.SinD(angleY)
    zr = sqr(temp1 + temp2)  

    dY = xr*CSTProject.SinD(angleX)
    dZ = xr*CSTProject.CosD(angleX)

    dX = yr*CSTProject.SinD(angleY)
    dZ = yr*CSTProject.CosD(angleY)

    angleFromPoint = CSTProject.ATn2D(dY, dX)
    xPoint = offSetX + zr*CSTProject.CosD(angleFromPoint + angleZ)
    yPoint = offSetY + zr*CSTProject.SinD(angleFromPoint + angleZ)
    zPoint = offSetZ + dZ

    ''Moves point onto face (not center of coax)
    if angleX = 0 And angleZ = 0 then
      yPoint = yPoint + (outerRadius + innerRadius)/2
    End if
    'elseif orientation = "y" then
    '  xr = length/2
    '  yr = xr*CSTProject.SinD(angleX)
    '  temp1 = xr*CSTProject.CosD(angleX)*xr*CSTProject.CosD(angleX)
    '  temp2 = yr*CSTProject.SinD(angleY)*yr*CSTProject.SinD(angleY)
    '  zr = sqr(temp1 + temp2)  
    'elseif orientation = "x" then
    '  xr = 0
    '  yr = length/2
    '  zr = yr*CSTProject.CosD(angleY)
    'End if
    With CSTProject.Pick
      .ClearAllPicks
      .PickFaceFromPoint componentName + ":" + solidName, xPoint, yPoint, zPoint
    End With
  End Sub

  Public Sub PickNegativeFace()
      xr = -length/2
      yr = xr*CSTProject.CosD(angleX)
      temp1 = xr*CSTProject.SinD(angleX)*xr*CSTProject.SinD(angleX)
      temp2 = yr*CSTProject.SinD(angleY)*yr*CSTProject.SinD(angleY)
      zr = sqr(temp1 + temp2)  

      dY = xr*CSTProject.SinD(angleX)
      dZ = xr*CSTProject.CosD(angleX)

      dX = yr*CSTProject.SinD(angleY)
      dZ = yr*CSTProject.CosD(angleY)

      angleFromPoint = CSTProject.ATn2D(dY, dX)
      xPoint = zr*CSTProject.CosD(angleFromPoint + angleZ)
      yPoint = zr*CSTProject.SinD(angleFromPoint + angleZ)
      zPoint = offSetZ + dZ
      With CSTProject.Pick
        .ClearAllPicks
        .PickFaceFromPoint componentName + ":" + coaxSolidName, xPoint, yPoint, zPoint
      End With
  End Sub

  Public Sub Origin(X, Y, Z)
    offsetX = X
    offsetY = Y
    offsetZ = Z
  End Sub

  Public Sub Create
    if isEmpty(InnerRadius) then
      InnerRadius = InnerFromOuter(charImpedance, OuterRadius)
    elseif isEmpty(OuterRadius) then
      OuterRadius = OuterFromInner(CharImpedance, InnerRadius)
    elseif isEmpty(charImpedance) then
      charImpedance = CalcImpedance
    End if
    OuterCylinder
    InnerCylinder

    With CSTProject.Solid 
      .Subtract componentName + ":" + solidName, componentName + ":" + innerName
    End With 

    Solid.MoveSolid Me, offSetX, offSetY, offSetZ, False
  End Sub

  Public Function CalcImpedance()
    charImpedance = 60*log((outerRadius)/(innerRadius))
    CalcImpedance = charImpedance
  End Function

  Public Function AddPort()
    With CSTProject.Port 
      .Reset 
      number = .StartPortNumberIteration() + 1
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
 
    With CSTProject.Transform 
      .Reset 
      .Name "port" + cStr(number)
      .Origin "Free" 
      .Center "0", "0", "0" 
      if orientation = "x" then
        .Angle "0", "90", "0"
      elseif orientation = "y" then
        .Angle "-90", "0", "0"
      end if
      .MultipleObjects "False" 
      .GroupObjects "False" 
      .Repetitions "1" 
      .MultipleSelection "False" 
      .Transform "Port", "Rotate" 
    End With 

    With CSTProject.Transform 
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
    With CSTProject.Cylinder 
      .Reset 
      innerName = CSTProject.Solid.getNextFreeName()
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
    With CSTProject.Cylinder 
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
