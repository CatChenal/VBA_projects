Attribute VB_Name = "MDL_Utils_VB_APIs"
Option Explicit
'================================================================================
'
' MDL_Utils_VB_APIs Aug-26-03 16:40
'
'================================================================================
Declare Function SetForegroundWindow Lib "user32" (ByVal Hwnd As Long) As Long
Declare Function IsIconic Lib "user32" (ByVal Hwnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long

  Const SW_NORMAL = 1     'Show window in normal size
  Const SW_MINIMIZE = 2   'Show window minimized
  Const SW_MAXIMIZE = 3   'Show window maximized
  Const SW_SHOW = 9       'Show window without changing window size

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias _
                        "SHGetPathFromIDListA" (ByVal pidl As Long, _
                                                ByVal pszPath As String) As Long

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
  "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, _
                  ByVal wParam As Long, lParam As Long) As Long

Const cstWinAppPath = "C:\WINNT\SYSTEM32\"
'=================================================================

Function ReturnOpenFileName(frmCaller As Form, strFileFilter As String, _
                                              Optional strInitialDir As String, _
                                              Optional strTitle As String) As String
  Dim OpenFile As OPENFILENAME
  Dim lReturn As Long
  
  OpenFile.lStructSize = Len(OpenFile)
  OpenFile.hwndOwner = frmCaller.Hwnd
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
    OpenFile.lpstrInitialDir = "U:\PublicDB\Fundamentals\"  '"S:\"
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
'
Function OpenCharMap()
  Call OpenApp("CHARMAP.EXE", cstWinAppPath)
End Function

Function OpenCalculator()
  Call OpenApp("CALC.EXE", cstWinAppPath)
End Function

Function OpenNotePad()
  Call OpenApp("NOTEPAD.EXE", cstWinAppPath)
End Function

Function OpenApp(strAppName As String, strAppPath As String)
  Dim Hwnd As Long, lngTemp As Long
  Dim varVal As Variant
  
  If Len(strAppName) = 0 Then Exit Function
  Hwnd = IsAppUp(strAppName)
  If Hwnd <> 0 Then
    lngTemp = SetForegroundWindow(Hwnd)
    lngTemp = ShowWindow(Hwnd, SW_NORMAL)
  Else
    varVal = Shell(strAppPath & strAppName, vbNormalFocus)
  End If
End Function

Function IsAppUp(strAppEXEName As String) As Variant
  Dim lpClassname As String
  Select Case strAppEXEName
    Case "CALC.EXE"
      lpClassname = "SciCalc"
    Case "NOTEPAD.EXE"
      lpClassname = "NOTEPAD"
    Case "SOL.EXE"
      lpClassname = "Solitaire"
    Case "WINHELP.EXE"
      lpClassname = "MW_WINHELP"
    Case "PBRUSH.EXE"
      lpClassname = "MSPaintApp"
    Case "Explorer.EXE"
      lpClassname = "ExploreWClass"
    Case "WORDPAD.EXE"
      lpClassname = "WordPadClass"
  End Select
  IsAppUp = FindWindow(lpClassname, vbNullString)
End Function
