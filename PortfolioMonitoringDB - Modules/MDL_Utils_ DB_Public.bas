Attribute VB_Name = "MDL_Utils_ DB_Public"
Option Compare Database
Option Explicit
'================================================================================
'
' MDL_Utils_DB_Public Oct-31-03 10:00
'
' Module contains procs that can be used in most Access DB.
' Note 1:   proc RelinkTblsToSharedBE requies FROM "frmRelink"
' Note 1.5: proc ReturnOpenFileName (in RelinkTblsToSharedBE) requires mdl
'           MDL_Utils_Vb_APIs
' Note 2: proc SaveToFile requires Scripting lib in refs
'
'================================================================================
Public Const cstAdmin = "cchenal"
Public strLoggedUser As String
'
'Stand-alone forms
Public Const cstFRM_Main = "frmMainTab"
Public Const cstFRM_Fields = "frmFields"
Public Const cstFRM_NewSer = "frmNewSeries"
Public Const cstFRM_Msg = "frmMsg"
Public Const cstFRM_Cal = "frmCalendar"
Public Const cstFRM_Closing = "frmClosingInfo"
Public Const cstFRM_Fund = "frmFundData"
'Subforms to main form:
Public Const cstSFRM_Folio = "sfrmPortfolio"
Public Const cstSFRM_Fore = "sfrmForecasts"
Public Const cstSFRM_Comp = "sfrmCompanies"
Public Const cstSFRM_Series = "sfrmSeries"    'uses flx grid control
Public Const cstSFRM_Summary = "sfrmSummary"  'uses flx grid control
Public Const cstSFRM_Analysis = "sfrmAnalysis"

Public Const cstBE_Name = "PPMBE.mdb"
Public Const cstBE_Folder = "S:\PublicDB\Fundamentals\Portfolio Monitoring\"

Const cstPropBE = "AppBackEnd"
Const cstPropTitle = "AppTitle"
Const cstPropIcon = "AppIcon"
Const cstAppTitle = "Portfolio Performance Monitor"
Const cstAppIcon = "Misc\GRAPH04.ICO"
                  
Function AutoExecProc()
  On Error GoTo AutoExec_Err
  DoCmd.RunCommand acCmdAppMaximize
  DoCmd.Hourglass True
  strLoggedUser = GetLoggedUser
  DoCmd.OpenForm cstFRM_Main, acNormal, "", "", , acNormal
  Call SetDBToolbars
  
AutoExec_Exit:
  DoCmd.Hourglass False
  Exit Function

AutoExec_Err:
  MsgBox "Error: " & Err.Number & " : " & Err.Description, vbExclamation, "AutoExec"
  Resume AutoExec_Exit
End Function

Public Function GetAdminInfo()
  Dim strMsg As String, strFile As String
  Dim intFileNum As Integer
  Dim lngOutcome As Long
 
  strMsg = UCase("administrator information") & vbCrLf & _
           "The db administrator name is stored in the public constant " & _
           "cstAdmin located in the module MDL_Utils_DB_Public." & _
           vbCrLf & "It is used to run several utility procedures when the db " & _
           "is opened on the admin's workstation." & vbCr & vbCrLf & "The current db " & _
           "administrator is " & cstAdmin & Chr(46)
  strFile = cstBE_Folder & "Misc\AdminInfo.txt"
  intFileNum = FreeFile
  Open strFile For Output As #intFileNum
  Print #intFileNum, strMsg
  Close #intFileNum
  
  lngOutcome = Shell(cstWinAppPath & "NOTEPAD.EXE " & strFile, vbNormalFocus)
  If lngOutcome = 0 Then
    MsgBox "A problem occurred opening Notepad." & vbCrLf & _
           "The Admin Info is located in:" & vbCrLf & strFile, vbExclamation, _
           "Open Admin Info File"
  End If
End Function

Function SetDBToolbars()
  DoCmd.ShowToolbar "Database", acToolbarNo
  DoCmd.ShowToolbar "Formatting (Form/Report)", acToolbarNo
  DoCmd.ShowToolbar "Formatting (Datasheet)", acToolbarNo
  DoCmd.ShowToolbar "Form View", acToolbarNo
  DoCmd.Maximize
End Function

Function SetStartupProperties() 'used when db recreated for delivery
  Call SetDBProperty(CurrentDb, cstPropTitle, dbText, cstAppTitle)
  Call SetDBProperty(CurrentDb, cstPropIcon, dbText, cstBE_Folder & cstAppIcon)
  Call SetDBProperty(CurrentDb, cstPropBE, dbText, cstBE_Folder & cstBE_Name)
  Call SetDBProperty(CurrentDb, "StartupShowDBWindow", dbBoolean, False)
  Call SetDBProperty(CurrentDb, "StartupShowStatusBar", dbBoolean, False)
End Function

Public Function IsLoaded(strForm As String) As Boolean
  Const cstFORM_DESIGN = 0
  IsLoaded = False
  If CurrentProject.AllForms(strForm).IsLoaded Then
    IsLoaded = (Forms(strForm).CurrentView <> cstFORM_DESIGN)
  End If
End Function

Public Sub RelinkTblsToSharedBE()
  Dim frm As Form
  Dim strNewCnx As String, strDBBE As String, strFileFilter As String
  strDBBE = "": strNewCnx = ""
  On Error GoTo RelinkTblsToSharedBEErr
  
  DoCmd.Hourglass True
  Set frm = Forms(cstFRM_Main)
  
  ' Get linked tables connection string: (assuming unique source)
  strDBBE = GetBackEndDB '=BE path: H:\Projects Active\Portfolio Monitoring\PPMBE.mdb
  If Len(strDBBE) = 0 Then Exit Sub 'db has no linked tbl
  
  If MsgBox("Do you wish to specify a path for the back end?", vbQuestion + vbYesNo, _
            "New Path?...") = vbYes Then
    strFileFilter = "Access DB" & Chr$(0) & "*.mdb" & Chr$(0) & "*.MDB"
    strNewCnx = ReturnOpenFileName(frm, strFileFilter, strDBBE)
  End If
  
  If (Len(strNewCnx) = 0) Or (strNewCnx <> strDBBE) Then 'no new path specified
    MsgBox "You have not selected a new path.", vbExclamation, "Relinking Operation Cancelled"
    GoTo RelinkTblsToSharedBEExit
  End If
      
  strNewCnx = ";DATABASE=" & strNewCnx
  Call RelinkCurrentDBTables(strNewCnx)
  
RelinkTblsToSharedBEExit:
  Set frm = Nothing
  DoCmd.Hourglass False
  Exit Sub
  
RelinkTblsToSharedBEErr:
  MsgBox "Error: " & Err.Number & " : " & Err.Description, vbExclamation, "RelinkTblsToSharedBE"
  Resume RelinkTblsToSharedBEExit
End Sub

Sub RelinkCurrentDBTables(strCnxString As String)
  Dim tdf As DAO.TableDef
  Dim strConnect As String
  
  If Len(strCnxString) = 0 Then Exit Sub
  For Each tdf In CurrentDb.TableDefs
    strConnect = vbNullString
    strConnect = tdf.Connect
    If Len(strConnect) > 0 Then  'table is linked
      tdf.Connect = strCnxString
      tdf.RefreshLink
    End If
    'Debug.Print tdf.Name & " relinked to " & tdf.Connect
  Next tdf
  Set tdf = Nothing
  CurrentDb.TableDefs.Refresh
  CurrentDb.Close
End Sub

Public Function CloseAllOpenFormsButCaller(frmCallingForm As Form)
  Dim frm As Form
  Dim str As String
  
  If MsgBox("Close all open forms?", vbQuestion + vbYesNo, "Close...") = vbYes Then
    ' Search for open AccessObject objects in AllForms collection.
    For Each frm In Application.CurrentProject.AllForms
      If frm.IsLoaded = True Then
        str = frm.Name
        If str <> frmCallingForm.Name Then DoCmd.Close acForm, str
      End If
    Next frm
  End If
  Set frm = Nothing
End Function

Public Function GetFileNameFromPath(strFullPath As String, Optional blnExcludeExtension As Boolean) As String
  Dim str As String
  blnExcludeExtension = (Not IsMissing(blnExcludeExtension)) And (blnExcludeExtension = True)
  str = StrReverse(strFullPath)
  str = StrReverse(Left(str, InStr(1, str, "\") - 1))
  If blnExcludeExtension Then str = Left(str, InStrRev(str, ".") - 1)
  GetFileNameFromPath = str
End Function

Public Function GetQtrFromDate(dteAny As Date) As Integer
  GetQtrFromDate = CInt(Format(dteAny, "q"))
End Function

Public Function GetGivenQtrDate(iQuarter As Integer, iYear As Integer, Optional GetStartDate = True) As Date
  Dim iMonth As Integer

  If IsMissing(GetStartDate) Or GetStartDate Then
    iMonth = (iQuarter * 3) - 2
    GetGivenQtrDate = DateSerial(iYear, iMonth, 1)     ': start date (default)
  Else
    iMonth = (iQuarter * 3)
    GetGivenQtrDate = DateSerial(iYear, iMonth + 1, 0) ': end date
  End If
End Function

Public Function GetPrevQtrEndDate(Optional dteAnyDate As Variant) As Date
  Dim intCurrentQ As Integer, prevQ As Integer
  Dim intPrevQYear As Integer, intPrevQMonth As Integer
  Dim PrevQEndDate As Date
  Dim dte As Date
  
  If Not IsMissing(dteAnyDate) And IsDate(dteAnyDate) Then
    dte = dteAnyDate
  Else
    dte = Date
  End If
  intCurrentQ = Format(dte, "Q")
  intPrevQYear = Year(dte) - Abs(intCurrentQ = 1)
  intPrevQMonth = (Abs(intCurrentQ = 1) * (12)) + intCurrentQ * 3 - 3
  PrevQEndDate = DateSerial(intPrevQYear, intPrevQMonth + 1, 0)
  GetPrevQtrEndDate = PrevQEndDate
End Function

Public Function GetBackEndDB() As String  'Assumes unique backend source
  Dim str As String
  Dim t As Integer
  str = "": t = 0
  'Initialize str:
  str = CurrentDb.TableDefs(t).Connect  'If first tbl is linked then, no loop
  Do Until Len(str) <> 0
    str = CurrentDb.TableDefs(t).Connect
    t = t + 1
  Loop
  GetBackEndDB = Nz(Mid$(str, 11), "") 'remove ;DATABASE=
End Function

Public Function GetBackEndDir() As String
  Dim strConnect As String
  strConnect = GetBackEndDB
  GetBackEndDir = Left(strConnect, InStrRev(strConnect, "\"))
End Function

Public Function GetDBTitleProp() As String
  GetDBTitleProp = CurrentDb.Properties(cstPropTitle).Value
  
GetDBTitlePropExit:
  Exit Function
  
GetDBTitlePropErr:
  If Err = 3270 Then  'property was not found.
    Err.Clear
    Call SetDBProperty(CurrentDb, cstPropTitle, dbText, cstAppTitle)
    Resume Next
  Else
    MsgBox "Err: " & Err.Number & vbCr & Err.Description, , "GetDBTitleProp"
    Resume GetDBTitlePropExit
  End If
End Function

Public Function GetDBBackEndProp() As String
  On Error Resume Next
  GetDBBackEndProp = CurrentDb.Properties(cstPropBE).Value
  If Err = 3270 Then  'property was not found.
    Err.Clear
    Call SetDBProperty(CurrentDb, cstPropBE, dbText, cstBE_Folder & cstBE_Name)
    Resume Next
  Else
    If Err <> 0 Then MsgBox "Err: " & Err.Number & vbCr & Err.Description, , "GetDBBackEndProp"
  End If
End Function

Public Function SetDBProperty(dbsTemp As Database, strName As String, _
                              intDataType As DAO.DataTypeEnum, varValue As Variant)
  Dim prpNew As DAO.Property
  ' Attempt to set the specified property.
  On Error GoTo Err_Property
  dbsTemp.Properties(strName) = varValue
  On Error GoTo 0

Exit_Property:
  Set prpNew = Nothing
  Exit Function

Err_Property:
  If Err = 3270 Then  'property was not found.
    Set prpNew = dbsTemp.CreateProperty(strName, intDataType, varValue)
    dbsTemp.Properties.Append prpNew
    Resume Next
  Else
    MsgBox "Error number: " & Err.Number & vbCr & Err.Description, , "SetDBProperty"
    Resume Exit_Property
  End If
End Function

Sub ListDBProperties()
  Dim prop As DAO.Property
  Dim lng As Long
  lng = 0
  Debug.Print "Name         Value "
  For Each prop In Application.CurrentDb.Properties
    lng = lng + 1
    On Error Resume Next
    Debug.Print lng & ": " & prop.Name & ":: " & prop.Value
    If Err <> 0 Then
      If Err = 3251 Then
        Err.Clear
        Resume Next
      Else
        MsgBox "Error (" & Err & "): " & Err.Description, vbExclamation, "Proc: ListDBProperties"
      End If
    End If
  Next prop
End Sub

Public Function SaveToFile(FileName$, Body$, blnOverWriteExisting As Boolean)
  Dim fso As Object
  Dim fsFile As Object
  On Error GoTo SaveToFileErr
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set fsFile = fso.CreateTextFile(FileName$, blnOverWriteExisting)
  fsFile.WriteLine (Body$)
  fsFile.Close
  
SaveToFileExit:
  Set fsFile = Nothing
  Set fso = Nothing
  Exit Function
SaveToFileErr:
  MsgBox "Error (" & Err & "): " & Err.Description, vbExclamation, "Proc: SaveToFile"
  Resume SaveToFileExit
End Function

