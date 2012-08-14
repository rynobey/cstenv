''Script for performing general operations on projects

''***** INITIALISATION *****

option explicit
dim wsEnv, shEnv, fsEnv, envPath, libPath

''Get VBScript objects
Set wsEnv = WScript
Set shEnv = CreateObject("WScript.Shell")
Set fsEnv = CreateObject("Scripting.FileSystemObject")

''Initialise variables
envPath = fsEnv.getAbsolutePathName(".") + "\"
libPath = envPath + "lib\"

''Enable external library inclusion
Execute fsEnv.OpenTextFile(libPath + "import.vbs", 1).ReadAll()
''Include external libaries
Execute include(libPath + "project").ReadAll()


''***** MAIN *****


finish

''***** HELPERS *****

sub finish()
  ''Release memory
  set wsEnv = nothing
  set shEnv = nothing
  set fsEnv = nothing
end sub

sub newProject(projName)
  createProject projName, false
end sub

'sub rmProject(projectName)
'end sub

'sub mvProject(source, destination)
'end sub
