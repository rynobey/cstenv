''This file provides utility functions for operations with CST MWS projects

function createProject(projName, returnProject)
  ''Function for creating a new CST MWS project and saving it
  'Returns the project if returnProject = true


  ''Initialise variables
  projPath = envPath + "projects\" + projName + "\"

  '''Enable external library inclusion
  'Execute fs.OpenTextFile(libPath + "import.vbs", 1).ReadAll()
  '''Include external libaries
  'Execute include(libPath + "dirs").ReadAll()

  'Create project folders if they do not yet exist
  CreateDirs(projPath + "model")
  CreateDirs(projPath + "simulations")
  CreateDirs(projPath + "results")
  CreateDirs(projPath + "cst")

  'Create and save a new CST MWS project if it does not yet exist
  set project = nothing
  if not existFileProject(projName) then
    set project = cst.newMWS
    project.SaveAs projPath + "cst\" + projName + ".cst" , False
    if not returnProject then
      project.Quit
      set project = nothing
    end if
  end if

  ''Set the return value
  set createProject = project
end function

function openProject(projName)
  ''Function for opening an existing project
  'Returns the project if opened, nothing if it does not exist

  ''Initialise variables
  projPath = envPath + "projects\" + projName + "\"

  ''Set the return value
  set openProject = nothing
  if existFileProject(projName) then
    set openProject = cst.OpenFile(projPath + "cst\" + projName + ".cst")
  end if
end function

function saveAndQuitProject(project)
  ''Function for saving and closing a project
  'Returns true if successfull, false otherwise

  ''Save the current state and quit
  'TODO: implement success check
  project.Save
  project.Quit

  ''Release memory
  set project = nothing
end function

function existFoldersProject(projName)
  ''Function to check if the project folders exist
  projPath = envPath + "projects\" + projName + "\"

  existFoldersProject = true
  with fs
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
end function

function existFileProject(projName)
  ''Function to check if the CST MWS project file exists
  projPath = envPath + "projects\" + projName + "\"

  if fs.FileExists(projPath + "cst\" + projName + ".cst") then
    existFileProject = true
  else 
    existFileProject = false
  end if
end function

function cleanProject(project)
  ''Function that cleans the project (deletes ALL changes)

	''Get the project path
	cstPath = project.getProjectPath("Project")
	cstRoot = project.getProjectPath("Root")
	projName = right(cstPath, len(cstPath) - inStrRev(cstPath, "\"))

  project.fileNew
  project.SaveAs cstRoot + "\" + projName + ".cst" , False
  set cleanProject = project
end function
