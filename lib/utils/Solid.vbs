Set obj = New Solid

Class Solid

  Private Project

  Public Sub Init(CSTProject)
    Set Project = CSTProject
  End Sub

  Public Sub RotateAndCopy(componentName, solidName, pivotX, pivotY, pivotZ, angleX, angleY, angleZ, repititions)
    With Project.Transform 
      .Reset 
      .Name componentName + ":" + solidName
      .Origin "Free" 
      .Center cStr(pivotX), cStr(pivotY), cStr(pivotZ)
      .Angle cStr(angleX), cStr(angleY), cStr(angleZ)
      .MultipleObjects "True" 
      .GroupObjects "True" 
      .Repetitions cStr(repititions)
      .MultipleSelection "False" 
      .Destination "" 
      .Material "" 
      .Transform "Shape", "Rotate" 
    End With 
  End Sub

  Public Sub MoveExternalSolid(componentName, solidName, offsetX, offsetY, offsetZ)
    With Project.Transform 
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

  Public Sub MoveSolid(sourceSolid, offsetX, offsetY, offsetZ, doUpdate)
    With Project.Transform 
      .Reset 
      .component sourceSolid.componentName
      .Name sourceSolid.componentName + ":" + sourceSolid.solidName 
      .Vector cStr(offsetX), cStr(offsetY), cStr(offsetZ) 
      .UsePickedPoints "False" 
      .InvertPickedPoints "False" 
      .MultipleObjects "False" 
      .GroupObjects "False" 
      .Repetitions "1" 
      .MultipleSelection "False" 
      .Transform "Shape", "Translate" 
    End With

    if doUpdate then
      sourceSolid.offSetX = sourceSolid.offSetX + offSetX
      sourceSolid.offSetY = sourceSolid.offSetY + offSetY
      sourceSolid.offSetZ = sourceSolid.offSetZ + offSetZ
    End if

  End Sub

  Public Sub RotateExternalSolid(componentName, solidName, pivotX, pivotY, pivotZ, angleX, angleY, angleZ)
    ''NOTE: Assumes rotation about only one axis!
    With Project.Transform 
      .Reset 
      .Name componentName + ":" + solidName 
      .Origin "Free" 
      .Center cStr(pivotX), cStr(pivotY), cStr(pivotZ)
      .Angle cStr(angleX), cStr(angleY), cStr(angleZ)
      .MultipleObjects "False" 
      .GroupObjects "False" 
      .Repetitions "1" 
      .MultipleSelection "False" 
      .Transform "Shape", "Rotate" 
    End With
  End Sub

  Public Sub RotateSolid(sourceSolid, pivotX, pivotY, pivotZ, angleX, angleY, angleZ, doUpdate)
    ''NOTE: Assumes rotation about only one axis!
    With Project.Transform 
      .Reset 
      .Name sourceSolid.componentName + ":" + sourceSolid.SolidName 
      .Origin "Free" 
      .Center cStr(pivotX), cStr(pivotY), cStr(pivotZ)
      .Angle cStr(angleX), cStr(angleY), cStr(angleZ)
      .MultipleObjects "False" 
      .GroupObjects "False" 
      .Repetitions "1" 
      .MultipleSelection "False" 
      .Transform "Shape", "Rotate" 
    End With

    if doUpdate then
      if angleX <> 0 then
      elseif angleY <> 0 then
        pivotLengthZ = (pivotZ - sourceSolid.offSetZ)
        pivotLengthX = (pivotX - sourceSolid.offSetX)
        pivotR = sqr(pivotLengthZ*pivotLengthZ + pivotLengthX*pivotLengthX)
        angleFromPivot = Project.atn2D(-pivotLengthX, -pivotLengthZ)
        endAngle = angleFromPivot + angleY
        sourceSolid.offSetZ = pivotZ + pivotR*Project.cosD(endAngle)
        sourceSolid.offSetX = pivotX + pivotR*Project.sinD(endAngle)
      elseif angleZ <> 0 then
        pivotLengthX = (pivotX - sourceSolid.offSetX)
        pivotLengthY = (pivotY - sourceSolid.offSetY)
        pivotR = sqr(pivotLengthX*pivotLengthX + pivotLengthY*pivotLengthY)
        angleFromPivot = Project.atn2D(-pivotLengthY, -pivotLengthX)
        endAngle = angleFromPivot + angleZ
        sourceSolid.offSetX = pivotX + pivotR*Project.cosD(endAngle)
        sourceSolid.offSetY = pivotY + pivotR*Project.sinD(endAngle)
      End if
      sourceSolid.AngleX = sourceSolid.AngleX + angleX
      sourceSolid.AngleY = sourceSolid.AngleY + angleY
      sourceSolid.AngleZ = sourceSolid.AngleZ + angleZ
    End if
  End Sub

  Public Sub AlignExternalSolid(componentName, solidName, orientation)
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
  End Sub

  Public Sub AlignSolid(sourceSolid, orientation, doUpdate)
    With Project.Transform 
      .Reset 
      .component sourceSolid.componentName
      .Name sourceSolid.componentName + ":" + sourceSolid.solidName 
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
      .Transform "Shape", "Rotate" 
    End With 

    if doUpdate then
      sourceSolid.orientation = orientation
      if orientation = "x" then
        tempZ = sourceSolid.offSetZ
        sourceSolid.offSetZ = -sourceSolid.offSetX
        sourceSolid.offSetX = tempZ
      elseif orientation = "y" then
        tempZ = sourceSolid.offSetZ
        sourceSolid.offSetZ = -sourceSolid.offSetY
        sourceSolid.offSetY = tempZ
      End if
    End if
  End Sub

  Public Sub LocalMeshSolid(componentName, solidName, dx, dy, dz, px, py, pz)
    Project.Solid.SetSolidLocalMeshProperties componentName + ":" + solidName, "PBA", "True", "0", "True", "False", cStr(dx), cStr(dy), cStr(dz), cStr(px), cStr(py), cStr(pz), "False", "1.0", "False", "1.0", "False", "True", "True", "True" 
  End Sub

End Class
