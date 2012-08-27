Set obj = New Project

Class Project

  Private FileSystem
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
    Set FileSystem = Env.Use("lib\utils\FileSystem")
  End Sub
  
  Public Sub InterpretCommand(cmd)
    if Trim(cmd) = "End" Or Trim(cmd) = "end" then
      Env.CurrentProject = ""
    else
      Execute cmd 
    End if
  End Sub

  Public Function Add(CSTProjectName, returnProject)
    ''Function for creating a new CST MWS project and saving it
    'Returns the project if returnProject = true
    Set temp = nothing
    if not FileSystem.CSTFileExists(projectName, CSTProjectName) then
      set temp = Env.CSTWrapper.CSTCOM.newMWS
      temp.SaveAs projectPath + "\cst\" + CSTProjectName + ".cst" , False
      if not returnProject then
        temp.Quit
        Set temp = Nothing
      else
        CSTProjects.Add CSTProjectName, temp
      End if
    End if
    Set Add = temp 
  End Function

  Public Function Open(CSTProjectName)
    Set temp = Nothing
    condition1 = FileSystem.CSTFileExists(projectName, CSTprojectName)
    condition2 = not CSTProjects.Exists(CSTProjectName)
    if condition1 And condition2 then
      Set temp = Env.CSTWrapper.CSTCOM.OpenFile(projectPath + "\cst\" + CSTProjectName + ".cst")
      CSTProjects.Add CSTProjectName, temp
    elseif not condition1 And condition2 then
      Set temp = Add(CSTProjectName, True)
      CSTProjects.Add CSTProjectName, temp
    elseif not condition2 then
      Env.Print("That CST Project is already open!")
    End if
    Set Open = temp
  End Function

  Public Function SetModel(CSTProjectName, CSTModelName)
    path = "projects\" + projectName + "\models\" + CSTModelName
    Set Temp = Env.Use(path)
    Temp.Init(CSTProjects.Item(CSTProjectName))
    if CSTModels.Exists(CSTProjectName) then
      Set CSTModels.Item(CSTProjectName) = Temp
    else
      CSTModels.Add CSTProjectName, Temp
    End if
    Set Temp = Nothing
  End Function

  Public Function Run(SimulationName, ResultsFolderName)
    path = "projects\" + projectName + "\simulations\" + SimulationName
    Set Sim = Env.Use(path)
    sim.Init(Me)
    sim.Run(ResultsFolderName)
    Set Sim = Nothing
  End Function

  Public Function Check(SimulationName)
    path = "projects\" + projectName + "\simulations\" + SimulationName
    Set Sim = Env.Use(path)
    sim.Init(Me)
    sim.Check()
    Set Sim = Nothing
  End Function

  Public Function Build(CSTProjectName, ParameterSet)
    Set model = CSTModels.Item(CSTProjectName)
    model.Build(ParameterSet)
    Set model = Nothing
  End Function

  Public Function DBuild(CSTProjectName)
    Set model = CSTModels.Item(CSTProjectName)
    model.DBuild()
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
