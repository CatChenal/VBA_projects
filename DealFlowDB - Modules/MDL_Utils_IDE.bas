Attribute VB_Name = "MDL_Utils_IDE"
Option Explicit
'================================================================================
'
' MDL_Utils_IDE Aug-26-03 16:40
'
'================================================================================
Enum ENUM_RefsUpdOptions
  cstLoadRefsToFile = 0
  cstLoadRefsFromFile = 1
End Enum
Public Const cstDBRefsFile = "RefsList.txt"

Function OutputModules(strSavePath As Variant)
' ?OutputModules(CurrentProject.Path & "\Modules")
'*******************************************************************************
' This function will export ALL modules in the currentdb as .txt and .bas files
'  to the given location if given else, "c:\" is the default location).
' If a file with the same name exists, there is an error.
'*******************************************************************************
  Dim proj As VBProject
  Dim mdl As Object 'Module
  Dim strModuleName As String, strNewLoc As String, strExt As String
  Dim i As Integer
  
  On Error GoTo OutputModules_Error
  If Len(strSavePath & "") = 0 Then
    strSavePath = "C:"
  Else
    If Right(strSavePath, 1) = "\" Then strSavePath = Left(strSavePath, Len(strSavePath) - 1)
  End If
  
  Debug.Print "Files exported to " & strSavePath & "\"
  Set proj = Application.VBE.ActiveVBProject
  For i = 1 To proj.VBComponents.Count '- 1
    If Not proj.VBComponents(i).CodeModule Is Nothing Then
      Set mdl = proj.VBComponents(i).CodeModule
      strModuleName = mdl.Name
      strNewLoc = strSavePath & "\" & strModuleName
      If Left(strModuleName, 5) = "Form_" Then
        strExt = ".cls"
      Else
        strExt = ".bas"
      End If
      proj.VBComponents(i).Export (strNewLoc & strExt)
      strNewLoc = strNewLoc & ".txt"
      DoCmd.OutputTo acOutputModule, strModuleName, acFormatTXT, strNewLoc, 0
      Debug.Print i & ": " & strModuleName & " exported in .txt and .cls or .bas format"
    End If
  Next i
  Debug.Print "End Output Modules"
  
Exit_OutputModules:
  Set mdl = Nothing
  Set proj = Nothing
  Exit Function
 
OutputModules_Error:
  MsgBox "Error: " & Err.Number & " - " & Err.Description, vbExclamation, "OutputModules Err"
  Resume Exit_OutputModules
End Function

Function GetReferencesList() As String
  Dim r As Integer
  Dim str As String
  str = ""
  For r = 1 To References.Count
    If r < References.Count Then
      str = str & References(r).FullPath & vbCrLf
    Else
      str = str & References(r).FullPath
    End If
  Next r
  GetReferencesList = str
End Function

Function LoadProjectRefs(strFileName As String, intLoadIntoWhat As ENUM_RefsUpdOptions)
'?LoadProjectRefs(CurrentProject.Path & "\RefsList.txt",1)
'
' ENUM_LoadRefs 0 = cstLoadRefsToFile:
' To save current project's references list to a flat file, one ref path per line,
' using a standard reference file tag (module level var: cstRefFileTag).
'
' ENUM_LoadRefs 1 = cstLoadRefsIntoProject:
' To load the current project with a list references from a flat file
' located using a standard reference file tag (module level var: cstRefFileTag).
'
  Dim ref As Reference
  Dim strLibName As String, strBody As String
  Dim intFileNum As Integer

  intFileNum = FreeFile
  If intLoadIntoWhat = cstLoadRefsToFile Then
    strBody = GetReferencesList
    Open strFileName For Output As #intFileNum
    Print #intFileNum, strBody
  Else
    Open strFileName For Input As #intFileNum      ' Create file name.
    
    Do While Not EOF(intFileNum)       ' Check for end of file.
      Set ref = Nothing
      Input #intFileNum, strLibName          ' Read line of data.
      If strLibName = Chr(32) Then Exit Do
      If strLibName = Chr(34) Then Exit Do
      
      On Error Resume Next
      Set ref = References.AddFromFile(strLibName)
      If Err = 0 Then
        Debug.Print "Loaded: " & strLibName
      Else
        If Err = 32813 Then
          Err.Clear
          Debug.Print "Preloaded: " & strLibName
          Resume Next
        Else
          GoTo LoadProjectRefsErr
        End If
      End If
    Loop
  End If
    
LoadProjectRefsExit:
  Close #intFileNum     ' Close file.
  Set ref = Nothing
  'Debug.Print "LoadProjectRefs over"
  Exit Function

LoadProjectRefsErr:
  MsgBox "Error: " & Err.Number & " - " & Err.Description, vbExclamation, "LoadProjectRefs Err"
  Resume LoadProjectRefsExit
End Function

