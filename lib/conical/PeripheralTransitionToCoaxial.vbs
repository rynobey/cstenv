Set obj = New PeripheralTransitionToCoaxial

Class PeripheralTransitionToCoaxial

  Public componentName
  Public portRadius
  Public numberOfPorts
  Public TemplateCoaxialTransmissionLine

  Private CSTProject
  Private Solid


  Public Sub Init(NewCSTProject)
    Set CSTProject = NewCSTProject
    Set Solid = Env.Use("lib\utils\Solid")
    Solid.Init(CSTProject)
  End Sub

  Public Sub AddToConicalTransmissionLine(ConicalTransmissionLine, NewNumberOfPorts, NewPortRadius, outerRadius, impedance)

    portRadius = NewPortRadius
    numberOfPorts = NewNumberOfPorts
    componentName = ConicalTransmissionLine.ComponentName
    theta1 = ConicalTransmissionLine.Theta1
    theta2 = ConicalTransmissionLine.Theta2
    dTheta = abs(theta1 - theta2)

    'Create coaxial tx line template
    Set CoaxialTransmissionLine = Env.Use("lib\coaxial\TransmissionLine")
    CoaxialTransmissionLine.Init(CSTProject)
    With CoaxialTransmissionLine
      .componentName = componentName
      coaxSolidName = CSTProject.Solid.getNextFreeName()
      .solidName = coaxSolidName
      .Origin portRadius, 0, (0.1 + (portRadius)*(CSTProject.Pi/180)*dTheta/2)/2
      .OuterRadius = outerRadius
      .CharImpedance = impedance
      .Length = 0.1 + (portRadius)*(CSTProject.Pi/180)*dTheta/2
      .Create
    End With

    'Rotate coaxial tx line template
    Solid.RotateSolid CoaxialTransmissionLine, 0, 0, 0, 0, 90-(theta2 + theta1)/2, 0, True

    With CSTProject.Transform 
      .Reset 
      .Name componentName + ":" + coaxSolidName
      .Origin "Free" 
      .Center "0", "0", "0" 
      .Angle "0", "0", cStr(360/numberOfPorts) 
      .MultipleObjects "True" 
      .GroupObjects "True" 
      .Repetitions cStr(numberOfPorts - 1)
      .MultipleSelection "False" 
      .Destination "" 
      .Material "" 
      .Transform "Shape", "Rotate" 
    End With 

    'Create the template pin
    pinRadius = CoaxialTransmissionLine.InnerRadius
    With CSTProject.Cylinder
      .Reset 
      pinSolidName = CSTProject.Solid.getNextFreeName()
      .Name pinSolidName 
      .Component componentName
      .Material ConicalTransmissionLine.Material 
      .OuterRadius cStr(pinRadius) 
      .InnerRadius "0.0" 
      .Axis "z" 
      ZStart = cStr(-2*(portRadius + pinRadius)*(CSTProject.Pi/180)*dTheta/2)
      ZStop = cStr(2*(portRadius + pinRadius)*(CSTProject.Pi/180)*dTheta/2)
      .Zrange ZStart, ZStop
      .Xcenter cStr(portRadius) 
      .Ycenter "0" 
      .Segments "0" 
      .Create 
    End With

    'Rotate the template pin
    With CSTProject.Transform 
      .Reset 
      .Name componentName + ":" + pinSolidName 
      .Origin "Free" 
      .Center "0", "0", "0" 
      .Angle "0", cStr(90-(theta2 + theta1)/2), "0" 
      .MultipleObjects "False" 
      .GroupObjects "False" 
      .Repetitions "1" 
      .MultipleSelection "False" 
      .Transform "Shape", "Rotate" 
    End With

    'Transform and copy the template pin
    With CSTProject.Transform 
      .Reset 
      .Name componentName + ":" + pinSolidName 
      .Origin "Free" 
      .Center "0", "0", "0" 
      .Angle "0", "0", cStr(360/numberOfPorts) 
      .MultipleObjects "True" 
      .GroupObjects "True" 
      .Repetitions cStr(numberOfPorts - 1) 
      .MultipleSelection "False" 
      .Destination "" 
      .Material "" 
      .Transform "Shape", "Rotate" 
    End With

    'Align with conical tx line
    offSetX = ConicalTransmissionLine.OffSetX
    offSetY = ConicalTransmissionLine.OffSetY
    offSetZ = ConicalTransmissionLine.OffSetZ
    Solid.AlignExternalSolid componentName, pinSolidName, orientation
    Solid.MoveExternalSolid componentName, pinSolidName, offsetX, offsetY, offsetZ

    'Subtract from conical line
    solid1Path = componentName + ":" + ConicalTransmissionLine.SolidName
    solid2Path = componentName + ":" + pinSolidName
    CSTProject.Solid.Subtract solid1Path, solid2Path

    'Add peripheral coaxial line to conical line
    solid1Path = componentName + ":" + ConicalTransmissionLine.SolidName
    solid2Path = componentName + ":" + CoaxialTransmissionLine.SolidName
    CSTProject.Solid.Add solid1Path, solid2Path

    'Update the solidPath of the template coaxial line
    CoaxialTransmissionLine.solidName = ConicalTransmissionLine.solidName

    'Save the reference to the template coaxial transmission line
    Set TemplateCoaxialTransmissionLine = CoaxialTransmissionLine

    'Add this object to the ConicalTransmissionLine as an addon
    ConicalTransmissionLine.Addons.Add "PeripheralTransitionToCoaxial", Me

  End Sub
  
End Class
