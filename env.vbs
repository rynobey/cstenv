''Script for performing general operations on projects

''***** INITIALISATION *****

option explicit
dim wsEnv, shEnv, fsEnv, stdout, stdin, envPath, libPath, msg, lineLength

''Get VBScript objects
set wsEnv = WScript
set shEnv = CreateObject("WScript.Shell")
set fsEnv = CreateObject("Scripting.FileSystemObject")
set stdout = wsEnv.stdout
set stdin = wsEnv.stdin

''Initialise variables
envPath = fsEnv.getAbsolutePathName(".") + "\"
libPath = envPath + "lib\"
lineLength = 76

msg = "Importing libraries"
showStatus false
''Enable external library inclusion
Execute fsEnv.OpenTextFile(libPath + "import.vbs", 1).ReadAll()
''Include external libaries
Execute include(libPath + "project").ReadAll()
showStatus true


''***** MAIN *****


finish

''***** HELPERS *****

function showStatus(completed)
  dim i
  if not completed = true then
    stdout.write(msg)
  else
    for i = 1 to lineLength-Len(msg) step 1
      stdout.write(".")
    next
    stdout.writeLine("DONE")
  end if
end function

sub finish()
  ''Release memory
  msg = "Exiting: Collecting garbage"
  showStatus false
  set wsEnv = nothing
  set shEnv = nothing
  set fsEnv = nothing
  set stdin = nothing
  showStatus true
  set stdout = nothing
end sub

sub newProject(projName)
  createProject projName, false
end sub

'sub rmProject(projectName)
'end sub

'sub mvProject(source, destination)
'end sub
