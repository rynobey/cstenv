''Ths file provides utility functions for folder operations

function CreateDirs(paths)
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
end function

function getAbsoluteParent(path)
	getAbsoluteParent = left(path, inStrRev(path, "\"))
end function
