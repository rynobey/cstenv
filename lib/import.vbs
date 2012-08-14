function include(pathToFile)
  set fsImport = CreateObject( "Scripting.FileSystemObject" )
  set include = fsImport.OpenTextFile(pathToFile + ".vbs", 1)
  ' Release the file system object
  set fsImport = nothing
end function
