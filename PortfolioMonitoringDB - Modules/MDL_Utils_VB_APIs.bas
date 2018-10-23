Attribute VB_Name = "MDL_Utils_VB_APIs"
Option Explicit
'================================================================================
'
' MDL_Utils_VB_APIs Oct-13-03 16:00
'
'
'================================================================================
Public Const cstWinAppPath = "C:\WINDOWS\"

Public Declare Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
                        "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Type OPENFILENAME
 lStructSize As Long
 hwndOwner As Long
 hInstance As Long
 lpstrFilter As String
 lpstrCustomFilter As String
 nMaxCustFilter As Long
 nFilterIndex As Long
 lpstrFile As String
 nMaxFile As Long
 lpstrFileTitle As String
 nMaxFileTitle As Long
 lpstrInitialDir As String
 lpstrTitle As String
 flags As Long
 nFileOffset As Integer
 nFileExtension As Integer
 lpstrDefExt As String
 lCustData As Long
 lpfnHook As Long
 lpTemplateName As String
End Type

Const WM_USER = 1024

Declare Function FindWindow Lib "user32" Alias _
  "FindWindowA" (ByVal lpClassname As String, ByVal lpWindowName As Long) As Long

Declare Function SendMessage Lib "user32" Alias _
  "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
                  ByVal wParam As Long, lParam As Long) As Long
'=================================================================
                  
Function ReturnOpenFileName(frmCaller As Form, strFileFilter As String, _
                                              Optional strInitialDir As String, _
                                              Optional strTitle As String) As String
  Dim OpenFile As OPENFILENAME
  Dim lReturn As Long
 
  OpenFile.lStructSize = Len(OpenFile)
  OpenFile.hwndOwner = frmCaller.hWnd
  OpenFile.hInstance = Application.hWndAccessApp
  OpenFile.lpstrFilter = strFileFilter    '"*.xls" & Chr(0) & "*.XLS"
  OpenFile.nFilterIndex = 1 '
  OpenFile.lpstrFile = String(257, 0)
  OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
  OpenFile.lpstrFileTitle = OpenFile.lpstrFile
  OpenFile.nMaxFileTitle = OpenFile.nMaxFile
  If (Not IsMissing(strInitialDir)) And (Len(strInitialDir) > 0) Then
    OpenFile.lpstrInitialDir = strInitialDir
  Else
    OpenFile.lpstrInitialDir = "S:\PublicDB\Fundamentals\"
  End If
  If IsMissing(strTitle) Then strTitle = "Locate File "
  OpenFile.lpstrTitle = strTitle
  OpenFile.flags = 0
  
  lReturn = GetOpenFileName(OpenFile)
  If lReturn = 0 Then
     ReturnOpenFileName = ""
  Else
     ReturnOpenFileName = Trim(OpenFile.lpstrFile)
  End If
End Function

Function AppDetect(strApplicationClassName As String) As Long
'This procedure detects if an app is running and registers it. '? AppDetect("IEFrame")
'strApplicationClassName:
'   Excel=  XLMAIN
'   Word=   OpusApp
'   iExplorer=  IEFrame
'-------------------------------------------------------------
  Dim hWnd As Long
  'This returns a valid handle number, or zero if app is not running
  hWnd = FindWindow(strApplicationClassName, 0)   '"XLMAIN", 0)
  AppDetect = hWnd
  If hWnd = 0 Then Exit Function       'app is not running
  SendMessage hWnd, WM_USER + 18, 0, 0   'app is running. Register it in the ROT
End Function

Function AppOpen(ByVal strAppAPIClassName As String, ByVal strAppObjectClassName As String, _
                     blnAppVisible As Boolean, blnAppWasRunning As Boolean) As Object
' strAppObjectClassName:  Excel: Excel.Application
'                         Word:  Word.Application
  Dim appObj As Object
  Dim strClass As String, strObject As String
  DoCmd.Hourglass True
  On Error GoTo AppOpenErr
  
  strClass = strAppAPIClassName
  strObject = strAppObjectClassName
  blnAppWasRunning = CBool(AppDetect(strAppAPIClassName))
  'Debug.Print "blnAppWasRunning: " & blnAppWasRunning
  
  If blnAppWasRunning Then
    Set appObj = GetObject(, strAppObjectClassName)
  Else
    'Debug.Print "App was not running"
    Set appObj = CreateObject(strAppObjectClassName)
  End If
  appObj.Visible = blnAppVisible
  Set AppOpen = appObj
 ' Debug.Print "AppOpen.Name: " & AppOpen.Name
  
AppOpenExit:
  DoCmd.Hourglass False
  Exit Function
  
AppOpenErr:
  MsgBox "Error (" & Err & "): " & Err.Description, vbExclamation, "Proc: AppOpen"
  Resume AppOpenExit
End Function

Function GetLoggedUser() As String ' Returns the network login name
' Used by AutoExec
  Dim lngPos As Long
  Dim strUserName As String
  Const cstStrBuffLen = 254
  On Error GoTo GetLoggedUserErr
  
  GetLoggedUser = "unknown"
  strUserName = String$(cstStrBuffLen, 0)
  If apiGetUserName(strUserName, cstStrBuffLen + 1) <> 0 Then
    strUserName = RTrim(Left$(strUserName, 8))
    lngPos = InStr(8, strUserName, vbNullString)
    Do While lngPos <> 0   'remove chr(0) '( trailing spaces)
      strUserName = Left(strUserName, Len(strUserName) - 1)
      lngPos = InStr(8, strUserName, Chr(0))
    Loop
  End If

GetLoggedUserExit:
  GetLoggedUser = strUserName
  Exit Function
  
GetLoggedUserErr:
  strUserName = "Error"
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "GetloggedUser"
  Resume GetLoggedUserExit
End Function
