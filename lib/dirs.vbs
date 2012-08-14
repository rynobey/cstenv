function CreateDirs(path)
' This subroutine creates multiple folders like CMD.EXE's internal MD command.
' By default VBScript can only create one level of folders at a time (blows
' up otherwise!).
'
' Argument:
' MyDirName   [string]   folder(s) to be created, single or
'                        multi level, absolute or relative,
'                        "d:\folder\subfolder" format or UNC

  set fsDirs = CreateObject( "Scripting.FileSystemObject" )
  ' Convert relative to absolute path
  strDirs = fsDirs.GetAbsolutePathName(path)
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
      strDirBuild = fsDirs.BuildPath(strDirBuild, arrDirs(i))
      if not fs.FolderExists(strDirBuild) then 
          fs.CreateFolder strDirBuild
      end if
  next
  ' Release the file system object
  set fsDirs = nothing
end function
