''Define the project name
projName = "test"

''Get VBScript objects
Set ws = WScript
Set sh = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
'Import the include function
Execute fs.OpenTextFile(libPath + "import.vbs", 1).ReadAll()

''Get CST MWS objects
set cst = CreateObject("CSTStudio.Application")
set proj = cst.newMWS

''Initialise variables
envPath = fs.getAbsolutePathName("..") + "\"
libPath = envPath + "lib\"
projPath = envPath + "projects\" + projName + "\"

''Include external libaries
Execute include(libPath + "dirs").ReadAll()

'Create project folders if they do not yet exist
CreateDirs(projPath + "model")
CreateDirs(projPath + "simulations")
CreateDirs(projPath + "results")

'Create and save a new CST MWS project if it does not yet exist
proj.SaveAs projPath + projName + ".cst" , False
proj.Quit

''Release memory
set cst = nothing
set sh = nothing
set ws = nothing
set fs = nothing
