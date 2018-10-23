Attribute VB_Name = "Util-OutputModules"
Option Compare Database
Option Explicit
'
'*****************************************************************
' DF Reports-Util-OutputModules Sep-24-02 10:35
'******************************************************************
'

Function OutputModules(strSavePath As Variant)
' ?OutputModules("H:\Projects Active\Deal Flow Database\Deal Flow Reports\Modules")
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
  MsgBox "Error Number: " & err.Number & " - Error Description :" & err.Description, _
            vbExclamation, "OutputModules Err"
  Resume Exit_OutputModules
End Function
