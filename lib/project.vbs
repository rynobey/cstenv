Set obj = New Project

Class Project

  Private dirs
  Public CSTProjects
  Public CSTModels
  Public projectPath
  Public projectName
  
  Private Sub Class_Initialize
    Set CSTProjects = CreateObject("Scripting.Dictionary")
    Set CSTModels = CreateObject("Scripting.Dictionary")
  End Sub

  Private Sub Class_Terminate
    for each key in CSTProjects.keys
      Close(key)
    Next
    Set CSTProjects = Nothing
    Set CSTModels = Nothing
  End Sub

  Public Sub Init(projName)
    projectName = projName
    projectPath = Env.envPath + "\projects\" + projectName
    Set dirs = Env.Use("lib\dirs")
  End Sub
  
  Public Sub InterpretCommand(cmd)
    if Trim(cmd) = "End" Or Trim(cmd) = "end" then
      Env.using = ""
    else
      Execute cmd 
    End if
  End Sub

  Public Sub Add(CSTProjectName)
    ''Function for creating a new CST MWS project and saving it
    'Returns the project if returnProject = true
    Set temp = nothing
    if not dirs.CSTFileExists(projectName, CSTProjectName) then
      set temp = Env.cst.CSTCOM.newMWS
      temp.SaveAs projectPath + "\cst\" + CSTProjectName + ".cst" , False
      if not returnProject then
        temp.Quit
        Set temp = Nothing
      else
        CSTProjects.Add CSTProjectName, temp
      End if
    End if
    Set AddCSTProject = temp 
  End Sub

  Public Function Open(CSTProjectName)
    Set temp = Nothing
    condition1 = dirs.CSTFileExists(projectName, CSTprojectName)
    condition2 = not CSTProjects.Exists(CSTProjectName)
    if condition1 And condition2 then
      set temp = Env.cst.CSTCOM.OpenFile(projectPath + "\cst\" + CSTProjectName + ".cst")
      CSTProjects.Add CSTProjectName, temp
    End if
    Set Open = temp
  End Function

  Public Function SetModel(CSTProjectName, CSTModelName)
    path = "projects\" + projectName + "\models\" + CSTModelName
    Set temp = Env.Use(path)
    temp.Init(CSTProjects.Item(CSTProjectName))
    if CSTModels.Exists(CSTProjectName) then
      Set CSTModels.Item(CSTProjectName) = temp
    else
      CSTModels.Add CSTProjectName, temp
    End if
    Set temp = Nothing
  End Function

  Public Function Run(CSTProjectName, SimulationName)
    path = "projects\" + projectName + "\simulations\" + SimulationName
    Set sim = Env.Use(path)

    Set script = Env.Use("lib\script")
    for each ParameterKey in sim.ParameterSetArray(0).Keys
      leftSide = "["
      firstTime = true
      for each ParameterSet in sim.ParameterSetArray
        if firstTime then
          firstTime = false
          leftSide = leftSide & ParameterSet.Item(ParameterKey)
        else
          leftSide = leftSide & ", " & ParameterSet.Item(ParameterKey)
        End if
      Next
      script.AddLine(ParameterKey & " = " & leftSide & "];")
    Next
    path = "projects\" + projectName + "\results\" + SimulationName + ".m"
    script.SaveAs(path)

    sim.Init(Me)
    'sim.Run()
    Set sim = Nothing
  End Function

  Public Function Build(CSTProjectName, ParameterSet)
    Set model = CSTModels.Item(CSTProjectName)
    model.Build(ParameterSet)
    Set model = Nothing
  End Function

  Public Function Mesh(CSTProjectName)
    Set model = CSTModels.Item(CSTProjectName)
    model.Mesh()
    Set model = Nothing
  End Function

  Public Function Clean(CSTProjectName)
    ''Function that cleans the project (deletes ALL changes)

    Set CSTProject = CSTProjects.Item(CSTProjectName)
    ''Get the project path
    cstPath = CSTProject.getProjectPath("Project")
    cstRoot = CSTProject.getProjectPath("Root")
    CSTProjectName = right(cstPath, len(cstPath) - inStrRev(cstPath, "\"))
    CSTProject.fileNew
    CSTProject.SaveAs cstRoot + "\" + CSTProjectName + ".cst" , False
    Set CSTProjects.Item(CSTProjectName) = CSTProject
  End Function

  Public Function Close(CSTProjectName)
    ''Save the current state and quit
    'TODO: implement success check
    if CSTProjects.Exists(CSTProjectName) then
      Set CSTProject = CSTProjects.Item(CSTProjectName)
      CSTProject.Save
      CSTProject.Quit
      CSTProjects.Remove(CSTProjectName)
    End if
    if CSTModels.Exists(CSTProjectName) then
      CSTModels.Remove(CSTProjectName)
    End if
    Set CSTProject = Nothing
  End Function

End Class
