Set env = New Environment
env.Main()
Set env = Nothing

Class Environment

  Private fs
  Private dirs
  Private EXIT_COMMAND
  Public using
  Public envPath
  Public libPath
  Public cst
  Public Projects
  
  Private Sub Class_Initialize()
    Set fs = CreateObject("Scripting.FileSystemObject")
    envPath = fs.getAbsolutePathName(".")
    Set Projects = CreateObject("Scripting.Dictionary")
    Set dirs = Use("lib\dirs")
    ExecuteGlobal "Set env = Me"
    EXIT_COMMAND = "exit"
    using = ""
    ConCST
  End Sub

  Private Sub Class_Terminate()
    Set cst = Nothing
    Set fs = Nothing
    Set dirs = Nothing
  End Sub

  Private Sub ConCST()
    if isEmpty(cst) then
      Set cst = Use("lib\cst")
    End if
  End Sub

  Private Sub InterpretCommand(cmd)
    'TODO: tokenize cmd str on ":" characters and loop over tokens
    if InStr(1, cmd, "Using", 1) = 1 then
      using = Trim(Right(cmd, Len(cmd) - Len("Using")))
    elseif using <> "" then
      Projects.Item(using).InterpretCommand(cmd)
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
      WScript.stdout.write(using & ">")
      cmd = WScript.stdin.ReadLine
      if cmd <> EXIT_COMMAND then
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
      WScript.stdout.write(using & ">")
      cmd = WScript.stdin.ReadLine
      if cmd <> EXIT_COMMAND then
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
    Set temp = Use("lib\project")
    temp.Init(projectName)
    Projects.Add projectName, temp
    Set OpenProject = temp
    Set temp = Nothing
  End Function

  Public Function CloseProject(projectName)
    Projects.Remove(projectName)
  End Function

  Public Function NewProject(projectName)
    'Create project folders if they do not yet exist
    path = envPath + "\projects\" + projectName + "\"
    dirs.CreateDirs(path + "models")
    dirs.CreateDirs(path + "simulations")
    dirs.CreateDirs(path + "results")
    dirs.CreateDirs(path + "cst")

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
