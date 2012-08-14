''Script for performing general operations on projects

''***** INITIALISATION *****

option explicit

''Define GLOBAL variables
dim ws, sh, fs, cst, stdout, stdin, envPath, libPath, msg, lineLength, project

''Get VBScript objects
set ws = WScript
set sh = CreateObject("WScript.Shell")
set fs = CreateObject("Scripting.FileSystemObject")

''Get CST MWS objects
set cst = CreateObject("CSTStudio.Application")

''Initialise variables
set stdout = ws.stdout
set stdin = ws.stdin
envPath = fs.getAbsolutePathName(".") + "\"
libPath = envPath + "lib\"
lineLength = 80

''Enable external library inclusion
Execute fs.OpenTextFile(libPath + "import.vbs", 1).ReadAll()

''Include external libaries
msg = "Importing libraries"
showStatus 0
Execute include(libPath + "project").ReadAll()
Execute include(libPath + "dirs").ReadAll()
showStatus 1


''***** MAIN *****

mainMenu
finish

''***** HELPERS *****

function showStatus(status)
	dim i
  if status = 0 then
    stdout.write(msg)
	elseif status = 1 then
		for i = 1 to lineLength-Len(msg)-Len("DONE") step 1
			stdout.write(".")
		next
		stdout.writeLine("DONE")
		stdout.writeLine("")
	elseif status = -1 then
		for i = 1 to lineLength-Len(msg)-Len("FAILED") step 1
			stdout.write(".")
		next
		stdout.writeLine("FAILED")
		stdout.writeLine("")
	end if
end function

sub mainMenu()
  dim entered, break, counter
  break = false
  while not break
    ''Show the available options
    stdout.writeLine("Choose an action:")
    stdout.writeLine("1) Open CST MWS Project")
    stdout.writeLine("2) Create New CST MWS Project")
    stdout.writeLine("3) Remove CST MWS Project")
    stdout.writeLine("4) Rename CST MWS Project")
    stdout.writeLine("Q) Quit")

    ''Indicate ready to receive input
		stdout.writeLine("")
    stdout.write(">")

    ''Handle input
    entered = stdin.readLine
    stdout.writeLine("")
    if entered = "1" then
      openProjectMenu
    elseif entered = "2" then
      newProjectMenu
    elseif entered = "3" then
      stdout.writeLine("3")
    elseif entered = "Q"  or entered = "q" then
      break = true
    end if
  wend
end sub

sub finish()
  ''Release memory
  msg = "Exiting: Collecting garbage"
  showStatus 0 
  set ws = nothing
  set sh = nothing
  set fs = nothing
  set stdin = nothing
  set cst = nothing
  showStatus 1 
  set stdout = nothing
end sub

sub openProjectMenu()
  dim entered, break
  break = false
  while not break
    ''Show the available options
    stdout.writeLine("1) Open project with name")
    stdout.writeLine("b) back")

    ''Indicate ready to receive input
    stdout.writeLine("")
    stdout.write(">")
    
    ''Handle input
    entered = stdin.readLine
    stdout.writeLine("")
    if entered = "b" then
      break = true
    elseif entered = "1" then
      ''Read name from stdin
      stdout.writeLine("Enter the project name:")

      ''Indicate ready to receive input
      stdout.writeLine("")
      stdout.write(">")
      
      ''Handle input
      entered = stdin.readLine
      stdout.writeLine("")
      if entered <> "!back" then
				msg = "Opening"
				showStatus 0
        set project = openProject(entered)
        if project is nothing then
					showStatus -1
        else
					showStatus 1
          projectMenu()
        end if
      end if
    end if
  wend
end sub

sub projectMenu()
	dim entered, break, cstPath, cstRoot, projName

	break = false

	''Get the project path
	cstPath = project.getProjectPath("Project")
	cstRoot = project.getProjectPath("Root")
	projName = right(cstPath, len(cstPath) - inStrRev(cstPath, "\"))

	while not break

		''Reload project scripts
		Execute fs.OpenTextFile(getAbsoluteParent(cstRoot) + "\main.vbs", 1).ReadAll()

		''Show the available options
		stdout.writeLine("1) Build model")
		stdout.writeLine("2) Save project")
    stdout.writeLine("3) Clean project (undo ALL changes)")
		stdout.writeLine("b) back (close project)")
		stdout.writeLine("")

		''Indicate ready to receive input
		stdout.write(projName + ">")

		''Handle input
		entered = stdin.readLine
		stdout.writeLine("")
		if entered = "b" then
			msg = "Saving project"
			showStatus 0
      project.SaveAs cstRoot + "\" + projName + ".cst" , False
			showStatus 1
			project.Quit
			set project = nothing
			break = true
		elseif entered = "1" then
			msg = "Building model"
			showStatus 0
			build(project)
			showStatus 1
		elseif entered = "2" then
			msg = "Saving project"
			showStatus 0
      project.SaveAs cstRoot + "\" + projName + ".cst" , False
			showStatus 1
    elseif entered = "3" then
      msg = "Cleaning (removing ALL changes)"
      showStatus 0
      set project = cleanProject(project)
      showStatus 1
		end if
	wend

end sub

sub newProjectMenu()
  dim entered, break
  break = false
  while not break
    ''Show the availble options
    stdout.writeLine("Enter the project name:")
    stdout.writeLine("")

    ''Indicate ready to receive input
    stdout.write(">")
    
    ''Handle input
    entered = stdin.readLine
    stdout.writeLine("")
    if entered = "!back" then
      break = true
    elseif entered <> "" then
      'TODO: implement name check for illegal characters
      msg = "Creating new project"
      showStatus 0
      createProject entered, false
      showStatus 1
      break = true
    end if
  wend
end sub

'sub rmProject(projectName)
'end sub

'sub mvProject(source, destination)
'end sub
