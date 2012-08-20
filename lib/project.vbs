Set obj = New Project

Class Project
''This class provides utility functions for operations with CST MWS projects

  Private dirs
  Public CSTProjects
  Public projectPath
  Public projectName
  
  Private Sub Class_Initialize
    Set CSTProjects = CreateObject("Scripting.Dictionary")
  End Sub

  Private Sub Class_Terminate
    Set CSTProjects = Nothing
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

    'Create and save a new CST MWS project if it does not yet exist
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

    ''Set the return value
    Set AddCSTProject = temp 
  End Sub

  Public Function Open(CSTProjectName)
    ''Function for opening an existing project
    'Returns the project if opened, nothing if it does not exist

    ''Set the return value
    Set temp = nothing
    condition1 = dirs.CSTFileExists(projectName, CSTprojectName)
    condition2 = not CSTProjects.Exists(CSTProjectName)
    if condition1 And condition2 then
      set temp = Env.cst.CSTCOM.OpenFile(projectPath + "\cst\" + CSTProjectName + ".cst")
      CSTProjects.Add CSTProjectName, temp
    End if
    Set Open = temp
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
    ''Function for saving and closing a project
    'Returns true if successfull, false otherwise

    ''Save the current state and quit
    'TODO: implement success check
    Set CSTProject = CSTProjects.Item(CSTProjectName)
    CSTProject.Save
    CSTProject.Quit

    ''Release memory
    CSTProjects.Remove(CSTProjectName)
    Set CSTProject = Nothing

  End Function

End Class
