Attribute VB_Name = "MDL_Outlook"
Option Compare Database
Option Explicit
'================================================================================
' MDL_Outlook Feb-18 10:25
'
'================================================================================
Public ol As Outlook.Application
Public ns As Outlook.NameSpace
Const cstMDL = "MDL Outlook"  'module contains functions to interface with MS Outlook
Const cstPublicFoldersDir = "Public Folders"
Const cstAllPublicFoldersDir = "All Public Folders"
'------------------------------------------------------------------------------------------

Function StartOutlook()
  On Error GoTo StartOutlookErr
  Set ol = New Outlook.Application
  Set ns = ol.GetNamespace("MAPI")                  'Reference the MAPI layer
  ns.Logon "MS Exchange Settings", , False, True    'New session with default settings, no dialog
  
StartOutlookExit:
  Exit Function
  
StartOutlookErr:
  Set ol = Nothing
  Set ns = Nothing
  MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, "StartOutlook"
  Resume StartOutlookExit
End Function
  
Function EndOutlook()
  If TypeName(ns) = "Nothing" Then
    Set ol = Nothing
    Exit Function
  End If
  ns.Logoff         'End current session
  Set ns = Nothing  'Release memory
  Set ol = Nothing
End Function

