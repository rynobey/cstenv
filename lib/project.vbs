''This file provides utility functions for operations with CST MWS projects

function createProject(projName, returnProject)
  ''Function for creating a new CST MWS project and saving it
  'Returns the project if returnProject = true

  ''Get VBScript objects
  Set wsProject = WScript
  Set shProject = CreateObject("WScript.Shell")
  Set fsProject = CreateObject("Scripting.FileSystemObject")

  ''Get CST MWS objects
  set cstProject = CreateObject("CSTStudio.Application")

  ''Initialise variables
  envPath = fsProject.getAbsolutePathName(".") + "\"
  libPath = envPath + "lib\"
  projPath = envPath + "projects\" + projName + "\"

  ''Enable external library inclusion
  Execute fsProject.OpenTextFile(libPath + "import.vbs", 1).ReadAll()
  ''Include external libaries
  Execute include(libPath + "dirs").ReadAll()

  'Create project folders if they do not yet exist
  CreateDirs(projPath + "model")
  CreateDirs(projPath + "simulations")
  CreateDirs(projPath + "results")
  CreateDirs(projPath + "cst")

  'Create and save a new CST MWS project if it does not yet exist
  set proj = nothing
  if not existFileProject(projName) then
    set proj = cstProject.newMWS
    proj.SaveAs projPath + "cst\" + projName + ".cst" , False
    if not returnProject then
      proj.Quit
      set proj = nothing
    end if
  elseif returnProject then
    set proj = cstProject.OpenFile(projPath + "cst\" + projName + ".cst")
  end if

  ''Release memory
  set cstProject = nothing
  set shProject = nothing
  set wsProject = nothing
  set fsProject = nothing

  ''Set the return value
  set createProject = proj
end function

function openProject(projName)
  ''Function for opening an existing project
  'Returns the project if opened, nothing if it does not exist

  ''Get VBScript objects
  Set fsProject = CreateObject("Scripting.FileSystemObject")

  ''Initialise variables
  envPath = fsProject.getAbsolutePathName("..") + "\"
  projPath = envPath + "projects\" + projName + "\"

  ''Get CST MWS objects
  set cstProject = CreateObject("CSTStudio.Application")

  ''Set the return value
  if existFileProject(projName) then
    set openProject = cstProject.OpenFile(projPath + "cst\" + projName + ".cst")
  end if

  ''Release memory
  set cstProject = nothing
  set fsProject = nothing
end function

function saveAndQuitProject(proj)
  ''Function for saving and closing a project
  'Returns true if successfull, false otherwise

  ''Save the current state and quit
  'TODO: implement success check
  proj.Save
  proj.Quit

  ''Release memory
  set proj = nothing
end function

function existFoldersProject(projName)
  ''Function to check if the project folders exist
  Set fsProject = CreateObject("Scripting.FileSystemObject")
  envPath = fsProject.getAbsolutePathName("..") + "\"
  projPath = envPath + "projects\" + projName + "\"

  existFoldersProject = true
  with fsProject
    if not .FolderExists(projPath) then
      existFoldersProject = false
    end if
    if not .FolderExists(projPath + "model") then
      existFoldersProject = false
    end if
    if not .FolderExists(projPath + "simulations") then
      existFoldersProject = false
    end if
    if not .FolderExists(projPath + "results") then
      existFoldersProject = false
    end if
    if not .FolderExists(projPath + "cst") then
      existFoldersProject = false
    end if
  end with

  ''Release memory
  set fsProject = nothing
end function

function existFileProject(projName)
  ''Function to check if the CST MWS project file exists
  Set fsProject = CreateObject("Scripting.FileSystemObject")
  envPath = fsProject.getAbsolutePathName("..") + "\"
  projPath = envPath + "projects\" + projName + "\"

  if fsProject.FileExists(projPath + "cst\" + projName + ".cst") then
    existFileProject = true
  else 
    existFileProject = false
  end if

  ''Release memory
  set fsProject = nothing
end function
