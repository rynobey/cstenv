Set obj = New Dirs

Class Dirs
''Ths class provides utility functions for folder operations
  Private fs

  Private Sub Class_Initialize
    Set fs = CreateObject("Scripting.FileSystemObject")
  End Sub

  Private Sub Class_Terminate
    Set fs = Nothing
  End Sub

  Public Function CreateDirs(paths)
    ' Argument:
    ' paths   [string]   folder(s) to be created, single or
    '                        multi level, absolute or relative,
    '                        "d:\folder\subfolder" format or UNC

    ' Convert relative to absolute path
    strDirs = fs.GetAbsolutePathName(paths)
    ' Split a multi level path in its "components"
    arrDirs = Split(strDirs, "\")
    ' Check if the absolute path is UNC or not
    if Left(strDirs, 2) = "\\" then
        strDirBuild = "\\" & arrDirs(2) & "\" & arrDirs(3) & "\"
        idxFirst    = 4
    else
        strDirBuild = arrDirs(0) & "\"
        idxFirst    = 1
    end if
    ' Check each (sub)folder and create it if it doesn't exist
    for i = idxFirst to Ubound(arrDirs)
        strDirBuild = fs.BuildPath(strDirBuild, arrDirs(i))
        if not fs.FolderExists(strDirBuild) then 
            fs.CreateFolder strDirBuild
        end if
    next
  End Function

  Public Function CSTFileExists(projectName, CSTProjectName)
    ''Function to check if the CST MWS project file exists
    path = Env.Projects.Item(projectName).ProjectPath + "\cst\"
    path = path + CSTProjectName + ".cst"
    if fs.FileExists(path) then
      CSTFileExists = true
    else 
      CSTFileExists = false
    end if
  End Function

  Function ProjectFoldersExist(projectName)
    projectPath = Env.envPath + "\projects\" + projectName
    ProjectFoldersExist = true
    With fs
      if not .FolderExists(projectPath) then
        ProjectFoldersExist = false
      end if
      if not .FolderExists(projectPath + "model") then
        ProjectFoldersExist = false
      end if
      if not .FolderExists(projectPath + "simulations") then
        ProjectFoldersExist = false
      end if
      if not .FolderExists(projectPath + "results") then
        ProjectFoldersExist = false
      end if
      if not .FolderExists(projectPath + "cst") then
        ProjectFoldersExist = false
      End if
    End With
  End Function

  Public Function GetAbsoluteParent(path)
    getAbsoluteParent = left(path, inStrRev(path, "\"))
  End Function

End Class
