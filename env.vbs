Set Env = New Environment
Env.Main()
Set Env = Nothing

Class Environment

  Private FS
  Private FileSystem
  Private EXIT_COMMAND
  Public currentProject
  Public envPath
  Public libPath
  Public CSTWrapper
  Public Projects
  
  Private Sub Class_Initialize()
    Set FS = CreateObject("Scripting.FileSystemObject")
    envPath = FS.getAbsolutePathName(".")
    Set Projects = CreateObject("Scripting.Dictionary")
    Set FileSystem = Use("lib\utils\FileSystem")
    ExecuteGlobal "Set Env = Me"
    EXIT_COMMAND = "exit"
    currentProject = ""
    ConCST
  End Sub

  Private Sub Class_Terminate()
    Set CSTWrapper = Nothing
    Set FS = Nothing
    Set FileSystem = Nothing
  End Sub

  Private Sub ConCST()
    if isEmpty(CSTWrapper) then
      Set CSTWrapper = Use("lib\utils\CSTWrapper")
    End if
  End Sub

  Private Sub Using(projectName)
    On Error Resume Next
    Err.Clear
    if Projects.Exists(projectName) then
      currentProject = projectName
    else
      OpenProject projectName
      currentProject = projectName
    End if
    if Err.Number <> 0 then
      Print("Failed to open project """ + projectName + """")
      currentProject = ""
      Err.Clear
    End if
  End Sub

  Private Sub InterpretCommand(cmd)
    if currentProject <> "" then
      Projects.Item(currentProject).InterpretCommand(cmd)
    else
      Execute cmd 
    End if
  End Sub

  Public Sub Print(msg)
    WScript.stdout.writeLine(msg)
  End Sub

  Public Sub Main()
    Set inputArguments = WScript.Arguments
    for each arg in inputArguments
      On Error Resume Next
      Err.Clear
      WScript.stdout.write(currentProject & ">")
      cmd = WScript.stdin.ReadLine
      if Trim(cmd) <> EXIT_COMMAND then
        InterpretCommand(cmd)
      End if
      if Err.Number <> 0 then 
        msg = "ERROR " & Err.Number & ": " & Err.Source
        msg = msg & ": " & Err.Description
        WScript.stdout.writeLine(msg)
        Err.Clear
      End if
    Next
    While cmd <> EXIT_COMMAND
      On Error Resume Next
      Err.Clear
      WScript.stdout.write(currentProject & ">")
      cmd = WScript.stdin.ReadLine
      if Trim(cmd) <> EXIT_COMMAND then
        InterpretCommand(cmd)
      End if
      if Err.Number <> 0 then 
        msg = "ERROR " & Err.Number & ": " & Err.Source
        msg = msg & ": " & Err.Description
        WScript.stdout.writeLine(msg)
        Err.Clear
      End if
    Wend
  End Sub

  Public Function OpenProject(projectName)
    Set Temp = Use("lib\utils\Project")
    Temp.Init(projectName)
    Projects.Add projectName, Temp
    Set OpenProject = Temp
    Set Temp = Nothing
  End Function

  Public Function CloseProject(projectName)
    Projects.Remove(projectName)
  End Function

  Public Function NewProject(projectName)
    'Create project folders if they do not yet exist
    path = envPath + "\projects\" + projectName + "\"
    FileSystem.CreateDirs(path + "models")
    FileSystem.CreateDirs(path + "simulations")
    FileSystem.CreateDirs(path + "results")
    FileSystem.CreateDirs(path + "cst")

    'TODO: Add creation of default scripts


    'Open and return the newly created project
    Set NewProject = OpenProject(projectName)
  End Function


  Public Function Use(file)
    location = envPath + "\" + file + ".vbs"
    Execute fs.OpenTextFile(location, 1).ReadAll()
    Set Use = obj
  End Function

End Class
