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

handleInput
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

sub showOptions()
  stdout.writeLine("")
  stdout.writeLine("Choose an action:")
  stdout.writeLine("1) Create New CST MWS Project")
  stdout.writeLine("2) Remove CST MWS Project")
  stdout.writeLine("3) Rename CST MWS Project")
  stdout.writeLine("Q) Quit")
  stdout.writeLine("")
end sub

sub handleInput()
  dim entered, break
  break = false

  do
    ''Show the available options
    showOptions
    ''Indicate ready for input
    stdout.write(">")

    ''Handle input
    entered = stdin.readLine
    if entered = "1" then
      stdout.writeLine("1")
    elseif entered = "2" then
      stdout.writeLine("2")
    else if entered = "Q" then
      break = true
    end if
  loop while not break
end sub

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
