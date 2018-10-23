Attribute VB_Name = "MDL_Utils_IDE"
Option Explicit
'================================================================================
'
' MDL_Utils_IDE Oct-21-03 9:50
'
'================================================================================
Enum ENUM_RefsUpdOptions
  cstLoadRefsToFile = 0
  cstLoadRefsFromFile = 1
End Enum
Const cstDBRefsFile = "RefsList.txt"

Function OutputModules(strSavePath As Variant)
' ?OutputModules(CurrentProject.Path & "\Modules")
'*******************************************************************************
' This function will export ALL modules in the currentdb as .txt and .bas files
'  to the given location if given else, "c:\" is the default location).
' If a file with the same name exists, there is an error. 12-04-03
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
  'Application.VBE.ActiveVBProject.VBComponents(1).Export("test.bas")

    If Not proj.VBComponents(i).CodeModule Is Nothing Then
      Set mdl = proj.VBComponents(i).CodeModule
      strModuleName = mdl.Name
      If Left(strModuleName, 5) = "Form_" Then
        strExt = ".cls"
      Else
        strExt = ".bas"
      End If
      strNewLoc = strSavePath & "\" & strModuleName & strExt
      proj.VBComponents(i).Export (strNewLoc)
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

Public Sub UpdateDBReferences(frmCaller As Form, intUpdateOption As ENUM_RefsUpdOptions)
  Dim strName As String, strPath As String, strFile As String, strTitle As String
  Const cstSaveRefs = "Save DB References"
  Const cstLoadRefs = "Load DB References"

  ' Get default location:
  strPath = Application.VBE.ActiveVBProject.FileName
  strName = GetFileNameFromPath(strPath)
  strPath = Mid$(strPath, 1, Len(strPath) - Len(strName)) '
  strTitle = Choose(intUpdateOption + 1, cstSaveRefs, cstLoadRefs) & _
                        Chr(40) & "Default name: " & cstDBRefsFile & Chr(41)
  strFile = ReturnOpenFileName(frmCaller, "*.txt", strPath, strTitle)
  If Len(strFile) = 0 Then Exit Sub
  Call LoadProjectRefs(strFile, intUpdateOption)
End Sub

Function GetReferencesList(Optional blnDetailedList As Boolean = False) As String
  Dim r As Integer, n As Integer
  Dim lngTilde As Long
  Dim str As String, strRef As String
  Const cstPROG = "PROGRA~"
  Const cstMICRO = "MICROS~"
  
  strRef = ""
  For r = 1 To References.Count
    str = References(r).FullPath
    
    ' Check for tilde in path which won't work on uploading:
    lngTilde = InStr(4, str, "~") 'start past drive ltr & first slash \
    Do While lngTilde <> 0
      n = Mid$(str, lngTilde + 1, 1)
      If InStr(4, str, cstPROG) Then
        str = Replace(str, cstPROG & n, "Program Files")
      ElseIf InStr(4, str, cstMICRO) Then
        str = Replace(str, cstMICRO & n, "Microsoft Office")
      End If
      lngTilde = InStr(4, str, "~")
    Loop
    If Not blnDetailedList Then
      strRef = strRef & str
    Else
      strRef = strRef & References(r).Name & ", " & str
    End If
    If r < References.Count Then strRef = strRef & vbCrLf
  Next r
  GetReferencesList = strRef
End Function

Function LoadProjectRefs(strFileName As String, intLoadIntoWhat As ENUM_RefsUpdOptions)
'?LoadProjectRefs(currentproject.path & "\RefsList.txt",1)
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
    Call CheckRefsListForCurrentUser(strFileName)
    
    Open strFileName For Input As #intFileNum
    Do While Not EOF(intFileNum)
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
  Debug.Print "LoadProjectRefs over"
  Exit Function

LoadProjectRefsErr:
  MsgBox "Error: " & Err.Number & " - " & Err.Description, vbExclamation, "LoadProjectRefs Err"
  Resume LoadProjectRefsExit
End Function

Function CheckRefsListForCurrentUser(strFileName As String) ''retired
' On first check, the mde ref will have the admin username (i.e. \\acnta030\cchenal$)
' Change it to current user, then rerun LoadProjRefs function
'
  Dim strLibName As String, strBody As String, strUser As String
  Dim lngDollar As Long, lngSlash As Long
  Dim intFileNum As Integer
  Dim blnSaveRefs As Boolean
  strBody = ""
  intFileNum = FreeFile
  
  Open strFileName For Input As #intFileNum
  Do While Not EOF(intFileNum)
    Input #intFileNum, strLibName          ' Read line of data.
    If strLibName = Chr(32) Then Exit Do
    If strLibName = Chr(34) Then Exit Do

    ' Check stored username on mde ref:
    lngDollar = InStr(3, strLibName, "$")  'start past \\
    If lngDollar <> 0 Then
      If strLoggedUser = "" Then strLoggedUser = GetLoggedUser
      If InStr(3, strLibName, strLoggedUser) = 0 Then 'not same user
        lngSlash = InStr(3, strLibName, "\")      'start past \\
        strUser = Mid$(strLibName, lngSlash + 1, lngDollar - lngSlash - 1)
        Debug.Print "Username in " & strLibName & " changed to: " & strLoggedUser
        blnSaveRefs = True
        strLibName = Replace(strLibName, strUser, strLoggedUser)
      Else
        Debug.Print "User in " & strLibName & " is logged user."
      End If
    End If
    strBody = strBody & strLibName & vbCrLf
  Loop
  Close #intFileNum     ' Close file.
  
  If blnSaveRefs Then Call vbSaveToFile(strFileName, strBody, True)

End Function

Private Function vbSaveToFile(FileName$, Body$, blnOverWriteExisting As Boolean)
  Dim intFileNum As Integer
  intFileNum = FreeFile
  If blnOverWriteExisting Then
    Kill FileName$
  Else
    Debug.Print FileName$ & " already exists and blnOverWriteExisting = " & blnOverWriteExisting
    Exit Function
  End If
  Open FileName$ For Output As #intFileNum
  Print #intFileNum, Body$
  Close #intFileNum     ' Close file.
End Function
