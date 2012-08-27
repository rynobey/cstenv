Set obj = New CenterTransitionToCoaxial

Class CenterTransitionToCoaxial

  ''Define public variables
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
  Public radius
  Public r2
  Public pathLength
  Public coaxialLengthRemoved
  Public conicalLengthRemoved


  ''Define private variables
  Private conical
  Private coaxial
  Private maxError
  Private bigTorusName
  Private smallTorusName
  Private cone1Name
  Private cone2Name
  Private cone3Name
  Private cylinderName
  Private CSTProject
  Private Solid

  Private Sub Class_Initialize
    'set up default values
    material = "Vacuum"
    componentName = "Transition"
    solidName = "CenterTransitionToCoaxial"
    offSetX = 0
    offSetY = 0
    offSetZ = 0
    angleX = 0
    angleY = 0
    angleZ = 0
  End Sub

  ''Define public methods
  Public Sub Init(NewCSTProject)
    Set CSTProject = NewCSTProject
    Set Solid = Env.Use("lib\utils\Solid")
    Solid.Init(CSTProject)
  End Sub

  Public Sub Origin(X, Y, Z)
    offsetX = X
    offsetY = Y
    offsetZ = Z
  End Sub

  Public Sub Link(conicalLine, coaxialLine, radius)
    Set conical = conicalLine
    Set coaxial = coaxialLine
    r1 = radius
    maxError = 0.05

    ''verify coaxial and conical impedances
    if abs(coaxial.CalcImpedance - conical.CalcImpedance) >= maxError then
      Env.Print("")
      Env.Print("Error: impedances are not equal: ")
      Env.Print(conical.CalcImpedance())
      Env.Print(" != ")
      Env.Print(coaxial.CalcImpedance())
    else
      ''determine the centres of the two tangent circles
      bR1 = coaxial.InnerRadius
      bR2 = coaxial.OuterRadius
      arg1 = (r1 - (bR2 + r1)*CSTProject.sinD(90-conical.Theta2))
      arg2 = (CSTProject.cosD(90-conical.Theta2))
      a = arg1/arg2
      arg1 = a*CSTProject.cosD(90-conical.Theta1)
      arg2 = bR1*CSTProject.sinD(90-conical.Theta1)
      arg1 = -(arg1 + arg2)
      arg2 = (CSTProject.sinD(90-conical.Theta1) - 1)
      r2 = arg1/arg2
       
      ''create profile
      CreateProfile bR1, bR2, r1, r2, a
      
      ''create transition
      CreateTransition bR1, bR2, r1, r2, a

      ''adjust conical tx line
      AdjustConicalLine bR1, bR2, r1, r2, a

      ''adjust coaxial tx line
      AdjustCoaxialLine bR1, bR2, r1, r2, a
    end if
  End Sub

  ''Define private methods

  Private Sub CreateProfile(bR1, bR2, r1, r2, a)
    With CSTProject.Torus 
      .Reset 
      .Component componentName
      bigTorusName = CSTProject.Solid.getNextFreeName()
      .Name bigTorusName
      .Material material 
      .OuterRadius cStr(bR1 + 2*r2)
      .InnerRadius cStr(bR1) 
      .Axis "z" 
      .Xcenter "0" 
      .Ycenter "0" 
      .Zcenter cStr(a) 
      .Segments "0" 
      .Create 
    End With 

    With CSTProject.Torus 
      .Reset 
      .Component componentName
      smallTorusName = CSTProject.Solid.getNextFreeName()
      .Name smallTorusName
      .Material material 
      .OuterRadius cStr(bR2 + 2*r1)
      .InnerRadius cStr(bR2) 
      .Axis "z" 
      .Xcenter "0" 
      .Ycenter "0" 
      .Zcenter cStr(a) 
      .Segments "0" 
      .Create 
    End With 
      
    arg1 = componentName + ":" + bigTorusName
    arg2 = componentName + ":" + smallTorusName
    CSTProject.Solid.Subtract arg1, arg2

    pathRadius = (bR1 + r2 + bR2 + r1)/2 - (coaxial.OuterRadius + coaxial.InnerRadius)/2
    pathAngle = CSTProject.Pi/180*(conical.Theta2 + abs(conical.Theta2 - conical.Theta1)/2)
    pathLength = pathRadius*pathAngle
  End Sub

  Private Sub CreateTransition(bR1, bR2, r1, r2, a)
    With CSTProject.Cone 
      .Reset 
      .Component componentName 
      cone1Name = CSTProject.Solid.getNextFreeName()
      .Name cone1Name
      .Material material
      .BottomRadius cStr(bR2 + r1 - r1*CSTProject.sinD(90-conical.Theta2))
      .TopRadius cStr(bR2 + r1)
      .Axis "z" 
      .Zrange cStr(a - r1*CSTProject.cosD(90-conical.Theta2)), cStr(a)
      .Xcenter "0" 
      .Ycenter "0" 
      .Segments "0" 
      .Create 
    End With 

    With CSTProject.Cone 
      .Reset 
      .Component componentName 
      .Name solidName
      .Material material
      .BottomRadius cStr(bR1 + r2 - r2*CSTProject.sinD(90-conical.Theta1))
      .TopRadius cStr(bR2 + r1 - r1*CSTProject.sinD(90-conical.Theta2))
      .Axis "z" 
      arg1 = cStr(a - r2*CSTProject.cosD(90-conical.Theta1))
      arg2 = cStr(a - r1*CSTProject.cosD(90-conical.Theta2))
      .Zrange arg1, arg2 
      .Xcenter "0" 
      .Ycenter "0" 
      .Segments "0" 
      .Create 
    End With 

    arg1 = componentName + ":" + solidName
    arg2 = componentName + ":" + cone1Name
    CSTProject.Solid.Add arg1, arg2 

    arg1 = componentName + ":" + solidName
    arg2 = componentName + ":" + bigTorusName
    CSTProject.Solid.Intersect arg1, arg2

    solid.MoveSolid Me, conical.offsetX, conical.offsetY, conical.offsetZ, True
  End Sub

  Private Sub AdjustConicalLine(bR1, bR2, r1, r2, a)
    With CSTProject.Cone 
      .Reset 
      .Component componentName 
      cone2Name = CSTProject.Solid.getNextFreeName()
      .Name cone2Name
      .Material material
      .BottomRadius cStr(bR2 + r1 - r1*CSTProject.sinD(90-conical.Theta2))
      .TopRadius cStr(bR2 + r1)
      .Axis "z" 
      if a > 0 then
        .Zrange cStr(a - r1*CSTProject.cosD(90-conical.Theta2)), cStr(a)
      else
        .Zrange cStr(a - r1*CSTProject.cosD(90-conical.Theta2)), "0"
      End if
      .Xcenter "0" 
      .Ycenter "0" 
      .Segments "0" 
      .Create 
    End With 

    With CSTProject.Cone 
      .Reset 
      .Component componentName 
      cone3Name = CSTProject.Solid.getNextFreeName()
      .Name cone3Name 
      .Material material
      .BottomRadius cStr(bR1 + r2 - r2*CSTProject.sinD(90-conical.Theta1))
      .TopRadius cStr(bR2 + r1 - r1*CSTProject.sinD(90-conical.Theta2))
      .Axis "z" 
      arg1 = cStr(a - r2*CSTProject.cosD(90-conical.Theta1))
      arg2 = cStr(a - r1*CSTProject.cosD(90-conical.Theta2))
      .Zrange arg1, arg2
      .Xcenter "0" 
      .Ycenter "0" 
      .Segments "0" 
      .Create 
    End With 

    arg1 = componentName + ":" + cone2Name
    arg2 = componentName + ":" + cone3Name
    CSTProject.Solid.Add arg1, arg2

    solid.MoveExternalSolid componentName, cone2Name, offsetX, offsetY, offsetZ

    arg1 = conical.componentName + ":" + conical.solidName
    arg2 = componentName + ":" + cone2Name
    CSTProject.Solid.Subtract arg1, arg2 

    arg1 = bR2 + r1 - r1*CSTProject.sinD(90-conical.Theta2)
    arg2 = bR1 + r2 - r2*CSTProject.sinD(90-conical.Theta1)
    ang = abs(conical.Theta2 + conical.Theta1)/2 - 90
    conicalLengthRemoved = (arg1 + arg2)/(2*CSTProject.cosD(ang))
  End Sub

  Private Sub AdjustCoaxialLine(bR1, bR2, r1, r2, a)
    solid.MoveSolid coaxial, offsetX, offsetY, offSetZ + a + coaxial.Length/2, True
  End Sub

End Class
