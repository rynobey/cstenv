Set obj = New TransmissionLine

Class TransmissionLine

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
  Public charImpedance
  Public Addons

  ''Define private variables
  Private profileName
  Private endName
  Private CSTProject
  Private Solid


  ''PUBLIC METHODS

  Public Sub Init(NewCSTProject)
    Set CSTProject = NewCSTProject
    Set Solid = Env.Use("lib\utils\Solid")
    Solid.Init(CSTProject)
    Set Addons = CreateObject("Scripting.Dictionary")
  End Sub

  Public Sub Create()
    if isEmpty(theta1) then 
      theta1 = CalcTheta1(charImpedance)
    End if
    CreateFace
    TransformFace
    RotateFace

    ''Remove temp objects
    CSTProject.Pick.ClearAllPicks
    CSTProject.Solid.Delete componentName + ":" + profileName

    SphericalEnd
    Solid.MoveSolid Me, offsetX, offsetY, offsetZ, False
  End Sub 

  Public Function CalcImpedance()
    charImpedance = 60*log((CSTProject.TanD(theta2/2))/(CSTProject.TanD(theta1/2)))
    CalcImpedance = charImpedance
  End Function

  Public Function CalcTheta1(impedance)
    CalcTheta1 = 2*CSTProject.AtnD((CSTProject.TanD(theta2/2))/(exp(impedance/60)))
  End Function


  ''PRIVATE METHODS

  Private Sub Class_Initialize
    'set up default values
    theta2 = 90
    charImpedance = 5
    offSetX = 0
    offSetY = 0
    offSetZ = 0
    angleX = 0
    angleY = 0
    angleZ = 0
    orientation = "z"
    material = "Vacuum"
    componentName = "Conical"
    solidName = "Tx"
  End Sub

  Private Sub CreateFace()
    dTheta = abs(theta1 - theta2)
    With CSTProject.AnalyticalFace
      .Reset 
      .Component componentName 
      profileName = CSTProject.Solid.getNextFreeName()
      .Name profileName
      .Material material
      .LawX "u"
      .LawY "0"
      .LawZ "v*u" + "*" + cStr(CSTProject.TanD(dTheta/2))
      .ParameterRangeU "0", radius
      .ParameterRangeV "-1", "1" 
      .Create
    End With
  End Sub

  Private Sub TransformFace()
    dTheta = abs(theta1 - theta2)
    With CSTProject.Transform 
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
    With CSTProject.Pick
      .ClearAllPicks
      .PickFaceFromId componentName + ":" + profileName, "1"
      .AddEdge "0.0", "0.0", "10", "0.0", "0.0", "-10"
    End With

    ''Rotate face to create conical transmission line
    With CSTProject.Rotate 
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
    With CSTProject.Sphere 
      .Reset 
      .Component componentName
      endName = CSTProject.Solid.getNextFreeName()
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
    CSTProject.Solid.Intersect arg1, arg2
  End Sub

End Class
