Set obj = New Results

Class Results

  Private CSTProject
  Private FileSystem

  Public Sub Init(NewCSTProject)
    Set CSTProject = NewCSTProject
    Set FileSystem = Env.Use("lib\utils\FileSystem")
  End Sub

  Private Function Get1DResultPaths()
    Set Paths = CreateObject("Scripting.Dictionary")
    root = "1D Results"
    CSTProject.SelectTreeItem(root)
    Set Get1DResultPaths = TraverseResultTree(root, Paths)
  End Function
  
  Private Function GetSParameterPaths()
    Set Paths = CreateObject("Scripting.Dictionary")
    root = "1D Results\S-Parameters"
    CSTProject.SelectTreeItem(root)
    Set GetSParameterPaths = TraverseResultTree(root, Paths)
  End Function

  Private Function GetPortSignalPaths()
    Set Paths = CreateObject("Scripting.Dictionary")
    root = "1D Results\Port signals"
    CSTProject.SelectTreeItem(root)
    Set GetPortSignalPaths = TraverseResultTree(root, Paths)
  End Function

  Public Function TraverseResultTree(Node, Dictionary)
    FirstChild = CSTProject.ResultTree.GetFirstChildName(Node)
    if FirstChild <> "" then ''This is a folder
      'loop over all children
      curChild = FirstChild
      while curChild <> ""
        Set Dictionary = TraverseResultTree(curChild, Dictionary)
        curChild = CSTProject.ResultTree.GetNextItemName(curChild)
      Wend
    else ''This is a file
      Dictionary.Add Node, 0
    End if
    Set TraverseResultTree = Dictionary
  End Function

  Public Function AppendParameterSet(ParameterSet, path)
      FileSystem.CreateDirs(path)
      text = ""
      for each k in ParameterSet.Keys
        v = ParameterSet.Item(k)
        text = text & vbNewLine & "if exist('" & k & "','var') == 0; " 
        text = text & k & " = " & v & "; else " & k & "(end+1) = " & v & "; end;"
      Next
      FileSystem.FileAppendLine path + "Parameters.m", text
  End Function

  Public Function AppendFrequency(path)
    FileSystem.CreateDirs(path)
    Set treePaths = GetSParameterPaths()
    for each treePath in treePaths
      dataFile = CSTProject.ResultTree.GetFileFromItemName(treePath)
      dataFile = Right(dataFile, Len(dataFile)-inStrRev(dataFile, "\"))
      dataFile = Left(dataFile, Len(dataFile)-Len(".sig"))
      Set Result = CSTProject.Result1DComplex(dataFile)
      arrX = Result1DXToCSV(Result)
      Exit For
    Next
    fileName = "Frequency.csv"
    FileSystem.FileAppendLine path & fileName, arrX
  End Function

  Public Function AppendTime(path)
    FileSystem.CreateDirs(path)
    Set treePaths = GetPortSignalPaths()
    for each treePath in treePaths
      dataFile = CSTProject.ResultTree.GetFileFromItemName(treePath)
      dataFile = Right(dataFile, Len(dataFile)-inStrRev(dataFile, "\"))
      dataFile = Left(dataFile, Len(dataFile)-Len(".sig"))
      Set Result = CSTProject.Result1D(dataFile)
      arrX = Result1DXToCSV(Result)
      Exit For
    Next
    fileName = "Time.csv"
    FileSystem.FileAppendLine path & fileName, arrX
  End Function

  Public Function AppendSParameters(path)
    path = path + "S-Parameters\"
    FileSystem.CreateDirs(path)
    Set treePaths = GetSParameterPaths()
    for each treePath in treePaths
      dataFile = CSTProject.ResultTree.GetFileFromItemName(treePath)
      dataFile = Right(dataFile, Len(dataFile)-inStrRev(dataFile, "\"))
      dataFile = Left(dataFile, Len(dataFile)-Len(".sig"))
      Set Result = CSTProject.Result1DComplex(dataFile)
      arrY = Result1DComplexYToCSV(Result)
      fileName = dataFile & ".csv"
      FileSystem.FileAppendLine path & fileName, arrY
    Next
  End Function

  Public Function AppendPortSignals(path)
    path = path + "Port Signals\"
    FileSystem.CreateDirs(path)
    Set treePaths = GetPortSignalPaths()
    for each treePath in treePaths
      dataFile = CSTProject.ResultTree.GetFileFromItemName(treePath)
      dataFile = Right(dataFile, Len(dataFile)-inStrRev(dataFile, "\"))
      dataFile = Left(dataFile, Len(dataFile)-Len(".sig"))
      Set Result = CSTProject.Result1D(dataFile)
      arrY = Result1DYToCSV(Result)
      fileName = dataFile & ".csv"
      FileSystem.FileAppendLine path & fileName, arrY
    Next
  End Function

  Public Function AppendBalance(path)
    FileSystem.CreateDirs(path)
    root = "1D Results\Balance"
    Set Paths = CreateObject("Scripting.Dictionary")
    Set treePaths = TraverseResultTree(root, Paths)
    for each treePath in treePaths
      dataFile = CSTProject.ResultTree.GetFileFromItemName(treePath)
      Set Result = CSTProject.Result1D(dataFile)
      arrY = Result1DYToCSV(Result)
      dataFile = Right(dataFile, Len(dataFile)-inStrRev(dataFile, "\"))
      dataFile = Left(dataFile, Len(dataFile)-Len(".bil"))
      fileName = "Balance" & dataFile & ".csv"
      FileSystem.FileAppendLine path & fileName, arrY
      Exit For
    Next
  End Function

  Public Function AppendEnergy(path)
    FileSystem.CreateDirs(path)
    root = "1D Results\Energy"
    Set Paths = CreateObject("Scripting.Dictionary")
    Set treePaths = TraverseResultTree(root, Paths)
    for each treePath in treePaths
      dataFile = CSTProject.ResultTree.GetFileFromItemName(treePath)
      Set Result = CSTProject.Result1D(dataFile)
      arrY = Result1DYToCSV(Result)
      dataFile = Right(dataFile, Len(dataFile)-inStrRev(dataFile, "\"))
      dataFile = Left(dataFile, Len(dataFile)-Len(".eng"))
      fileName = "Energy" & dataFile & ".csv"
      FileSystem.FileAppendLine path & fileName, arrY
    Next
  End Function

  Private Function Result1DComplexYToCSV(Result)
    arrY = ""
    N = Result.GetN
    for index = 0 To N-1 Step 1
      YRe = Result.getYRe(index)
      YIm = Result.getYIm(index)
      if index = 0 then
        if YIm > 0 then
          arrY = arrY & cStr(YRe) & "+" & cStr(YIm) & "j"
        else
          arrY = arrY & cStr(YRe) & cStr(YIm) & "j"
        End if
      else
        if YIm > 0 then
          arrY = arrY & "," & cStr(YRe) & "+" & cStr(YIm) & "j"
        else
          arrY = arrY & "," & cStr(YRe) & cStr(YIm) & "j"
        End if
      End if
    Next
    Result1DComplexYToCSV = arrY
  End Function

  Private Function Result1DYToCSV(Result)
    arrY = ""
    N = Result.GetN
    for index = 0 To N-1 Step 1
      if index = 0 then
        arrY = arrY & cStr(Result.getY(index))
      else
        arrY = arrY & "," & cStr(Result.getY(index))
      End if
    Next
    Result1DYToCSV = arrY
  End Function

  Private Function Result1DXToCSV(Result)
    arrX = ""
    N = Result.GetN
    for index = 0 To N-1 Step 1
      if index = 0 then
        arrX = arrX & cStr(Result.getX(index))
      else
        arrX = arrX & "," & cStr(Result.getX(index))
      End if
    Next
    Result1DXToCSV = arrX
  End Function

End Class
