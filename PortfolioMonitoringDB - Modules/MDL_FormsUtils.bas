Attribute VB_Name = "MDL_FormsUtils"
Option Compare Database
Option Explicit
'================================================================================
'
' MDL FormsUtils: Mar-16-04 11:30
' Oct-21-03 10:45
' Moved public refs for Main, Relink and Fields forms to
' MDL_Utils_DB_Public where there are first referenced (in AutoExec proc).
'
'================================================================================

Public blnNewCo As Boolean, blnFieldsUpdate As Boolean
Public blnExcelAlreadyRunning As Boolean    ' Flag for final release. set by GetPFMExcelBook
Public blnWordAlreadyRunning As Boolean     ' Flag for final release. set by OpenPFMReview
'
Public blnNoFrc As Boolean, blnNoPoints As Boolean
Public blnNewClosing As Boolean
'
Public lngPrevCo As Long, lngPrevFrc As Long
Public lngCurrentForecast As Long   ' Set every time RefreshGrid is called
Public lngCurrentComp  As Long
Public lngC As Long, lngR As Long   ' Col & Row num: Reset after the grids are refreshed.
'
Public Const cstQryCoFlds = "qryCoFieldsList"
'
Public Const cstAAColor = 5785120
Public Const cstGrey = 12632256

Public blnCopyField As Boolean
Public blnAddForecast As Boolean
Public strFieldsFormFilter As String
'
Public Const cstNoFrcForThisComp = "No forecast setup for this company"
'
'Const cstBTN_LastestFin = "lblBtn0"
Const cstBTN_ForecastsList = "lblBtn1"
'Const cstBTN_ForecastData = "lblBtn2"
Const cstBTN_SummaryForm = "lblBtn3"
'Const cstBTN_AnalysisReviews = "lblBtn4"
'Const cstBTN_Companies = "lblBtn5"
'
Const cstMDL = "FormsUtils"

Function CheckButtonSelection(strCallerCtl$) As Integer
' Outcome: 1 = success; 0 = failure
' For pages that need a forecast id
' Input: strCallerCtl$= name of the label that is clicked.

  Dim strMsg As String, strTitle As String
  Dim iOutcome
  strMsg = "": strTitle = "": iOutcome = 1
  
  Call HideMsgForm
  If Len(Trim(strCallerCtl$)) = 0 Then Exit Function  ' wrong input
  '
  If strCallerCtl$ = cstBTN_SummaryForm Then 'Summary page
    If Not CoHasDefaultBudget(Forms(cstFRM_Main)!cbxSelComp) Then
      iOutcome = 0
      strMsg = "Please set a default forecast before using the Summary Form."
      strTitle = "No default forecast."
    Else
      If Not blnNoFrc And (Forms(cstFRM_Main)!cbxSelForecast.Column(3) <> 0) Then
        iOutcome = 0
        strMsg = "You have selected the default budget forecast" & _
                  vbCrLf & "...to be compared with itself."
        strTitle = "Please select another forecast."
      End If
    End If
  Else
    If blnNoFrc Then
      strMsg = cstNoFrcForThisComp & ": " & Forms(cstFRM_Main)!cbxSelComp.Column(1)
      strTitle = "No forecasts"
      iOutcome = 0
      If strCallerCtl$ = cstBTN_ForecastsList Then
        strMsg = strMsg & vbCrLf & "Do you want to create a forecast now?"
        iOutcome = 1
        If MsgBox(strMsg, vbQuestion + vbYesNo, strTitle) = vbYes Then
          blnAddForecast = True
        Else
          blnAddForecast = False
        End If
      End If
    End If
  End If

  If iOutcome = 0 Then
    MsgBox strMsg, vbExclamation, strTitle
    Call RevertMainPrevBtn
    Call ToggleMainColor
  End If
  strMsg = ""
  CheckButtonSelection = iOutcome
End Function

Public Sub AddForecast()
  If Forms!frmMainTab!sfrmAny.Form.Name <> cstSFRM_Fore Then Exit Sub
  With Forms!frmMainTab!sfrmAny.Form
    .DataEntry = True
    .AllowAdditions = True
    .AllowEdits = True
    !cbxCompID.SetFocus
    !cbxCompID = lngCurrentComp
  End With
End Sub

Function ResetSubForm1Frcst()
  On Error GoTo ResetSubForm1Frcst_Err
  If (CheckButtonSelection(cstBTN_ForecastsList) <> 0) Then Call ResetSubForm(cstSFRM_Fore)
  Exit Function

ResetSubForm1Frcst_Err:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , cstMDL & ": ResetSubForm1Frcst"
End Function

Sub RevertMainPrevBtn()
' Reset the controls used to implement the color change to a previous state (cancel action).
  Dim strTemp As String
  With Forms(cstFRM_Main)
    strTemp = !txtCurBtn
    !txtCurBtn = !txtPrevBtn
    !txtPrevBtn = strTemp
  End With
End Sub

Function ToggleMainColor()
  Call ToggleCtlColors(Forms(cstFRM_Main), Forms(cstFRM_Main)!txtPrevBtn)
  Call ToggleCtlColors(Forms(cstFRM_Main), Forms(cstFRM_Main)!txtCurBtn)
End Function

Function ChangeAndToggleMainBtn(iBtnIdx As Integer)
  Dim strLbl As String

  strLbl = "lblBtn" & iBtnIdx 'becomes the name of the label control for referencing
  If Forms(cstFRM_Main)!txtCurBtn = strLbl Then Exit Function  'same btn clicked
  With Forms(cstFRM_Main)
    !txtPrevBtn = !txtCurBtn
    !txtCurBtn = strLbl
  End With
  Call ToggleMainColor
End Function

Function ResetSubForm(strSub$)
  DoCmd.Hourglass True
  With Forms(cstFRM_Main)
    ' Reload subform if Analysis
    If strSub$ = cstSFRM_Analysis Then
      If !cbxSelComp.RowSource <> "qrySelCompID_UR" Then
        !cbxSelComp.RowSource = "qrySelCompID_UR" 'List unrealized deals only
      End If
      !sfrmAny.SourceObject = strSub$
    Else
      If !sfrmAny.SourceObject = strSub$ Then Exit Function
      !cbxSelComp.RowSource = "qrySelCompID"  'List all companies
      !sfrmAny.SourceObject = strSub$
    End If
  End With
  Call MainFormSelComp_AfterUpdate
  DoCmd.Hourglass False
End Function

Sub ResetClosingOrFundForm(strFrm As String, blnNew As Boolean)
  Dim c As Integer
  On Error GoTo ResetClosingOrFundFormErr
  
  With Forms(strFrm)
    .AllowAdditions = blnNew
    .DataEntry = blnNew
    !lblNew.Visible = Not blnNew
     c = Abs((blnNew Or .Dirty))
    .Cycle = c
    !lblSave.Caption = Choose(c + 1, "Close", "Save")
    
    Select Case strFrm
      Case cstFRM_Closing
        !lblAlloc.Visible = Not .Dirty
        !cbxCompID = lngCurrentComp
        
      Case cstFRM_Fund
        !cbxClosingID = Forms(cstFRM_Closing)!lngClosingID
    End Select
    If Not blnNew Then .Requery
  End With
  
ResetClosingOrFundFormExit:
  Exit Sub
ResetClosingOrFundFormErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , cstMDL & ": ResetClosingOrFundForm"
  Resume ResetClosingOrFundFormExit
End Sub

Function HideMsgForm()
  If IsLoaded(cstFRM_Msg) Then Forms(cstFRM_Msg).Visible = False
End Function

Function SyncSubRecord(strFieldToSync1 As String, varValueToSync1 As Variant, _
                        Optional strFieldToSync2 As String = "", _
                        Optional varValueToSync2 As Variant = Null) As Integer
' Synchronized main form & subform record with at most two cbx values (AND criterion).
' Returned value: 0=OK; 1=nomatch, 9=other error
  Dim sfrm As Form
  Dim rst As DAO.Recordset
  Dim strCrit$
  Dim iOutcome As Integer
  iOutcome = 0 'ok
  On Error GoTo ProcErr
  
  If IsNull(varValueToSync1) Then Exit Function
  If Forms(cstFRM_Main)!sfrmAny.Form Is Nothing Then Exit Function
  
  ' Start forming the criteria for bookmarking:
  strCrit$ = "[" & strFieldToSync1 & "]=" & varValueToSync1
  If (Len(strFieldToSync2) > 0) And Not IsNull(varValueToSync2) Then
    If varValueToSync2 <> 0 Then
      strCrit$ = strCrit$ & " AND [" & strFieldToSync2 & "]=" & varValueToSync2
    End If
  End If
  
  If blnAddForecast Then
    Call AddForecast
  Else
    Set sfrm = Forms(cstFRM_Main)!sfrmAny.Form
    Set rst = sfrm.RecordsetClone
    rst.FindFirst strCrit$
    If rst.NoMatch Then
      iOutcome = 1 'reset
      GoTo ProcExit
    End If
    Forms(cstFRM_Main)!sfrmAny.Form.Bookmark = rst.Bookmark
  End If
  
ProcExit:
  rst.Close
  SyncSubRecord = iOutcome
  Set sfrm = Nothing
  Set rst = Nothing
  Exit Function

ProcErr:
  If Err <> 2467 Then 'no subform set
    iOutcome = 9
    MsgBox "Error: " & Err & ", " & Err.Description, vbExclamation, cstMDL & ": SyncSubRecord"
  End If
  Resume ProcExit
End Function

Function ToggleCtlColors(frm As Form, strCtlName As String)
' Implements the lbl color change when clicked to mimick a btn behavior.
' Toggles the background/foreground color of a ctl
  Dim lngBackColor As Long
  Dim ctl As Control
  'On Error GoTo ToggleCtlColorsErr
  
  If strCtlName >= " " Then
    Set ctl = frm(strCtlName)
    With ctl
      lngBackColor = .BackColor
      .BackColor = .ForeColor
      .ForeColor = lngBackColor
    End With
  End If
  
ToggleCtlColorsExit:
  Set ctl = Nothing
  Set frm = Nothing
  Exit Function
  
ToggleCtlColorsErr:
  MsgBox "Error: " & Err.Number & ": " & Err.Description, vbExclamation, cstMDL & ": ToggleCtlColors"
  Resume ToggleCtlColorsExit
End Function

Function Validate()
  On Error GoTo ClickValidateErr
  Application.Screen.MousePointer = 11 'busy
  
  Call ValidateFields(lngCurrentComp)
  
  If IsLoaded(cstFRM_Fields) Then
    Forms(cstFRM_Fields)!cbxCompID.Requery
    Forms(cstFRM_Fields).Requery
  End If
  Application.Screen.MousePointer = 0
  
ClickValidateExit:
  Exit Function
  
ClickValidateErr:
  Application.Screen.MousePointer = 0
  MsgBox "Error: " & Err.Number & ", " & Err.Description, vbExclamation, cstMDL & ": Validate"
  Resume ClickValidateExit
End Function

Sub ValidateFields(lngCoID As Long)
' Are ALL the required fields defined for the frcst(company)?: if not add them.
'
  Dim dbs As DAO.Database
  Dim rstReqFlds As DAO.Recordset, rstCoFlds As DAO.Recordset
  Dim str$, strAdded$, strName$
  Dim r As Integer, iResult As Integer
  Dim varPos, varCoReqFields()
  Const cstQryReqFlds = "qryRequiredFields"    'as defined in the lookup tbl
  Const cstQryCoReqFlds = "qryCoReqFieldsList"
  r = 0: strAdded$ = ""
  
  On Error GoTo ValidateFieldsErr
      
  Set dbs = CurrentDb
   
  iResult = GetCoFieldsArray(lngCoID, cstQryCoReqFlds, varCoReqFields)
  ' Next, the processing will determine if the company's required fields match
  ' those in the lookup table:
  If iResult <> 0 Then 'err
    GoTo ValidateFieldsExit
  Else
  
    ' Open the req flds qry to iterate and compare with the set required fields in company
    Set rstReqFlds = dbs.OpenRecordset(cstQryReqFlds) ': the 'official' list
    rstReqFlds.MoveLast
    rstReqFlds.MoveFirst
    
    ' Open tbl in case addition/update needed:
    Set rstCoFlds = dbs.TableDefs("tblCompFields").OpenRecordset(dbOpenDynaset)
    rstCoFlds.MoveLast
  
    For r = 0 To rstReqFlds.RecordCount - 1
      str$ = rstReqFlds!txtDispName
      
      If Not IsInArray(varCoReqFields, str$, True) Then
        rstCoFlds.AddNew
        rstCoFlds!lngCompID = lngCoID
        rstCoFlds!lngFldID = rstReqFlds!lngFldID
        rstCoFlds!lngAcctgTypeID = rstReqFlds!lngAcctgTypeID
        rstCoFlds!lngPriorityID = rstReqFlds!lngPriorityID
        rstCoFlds.Update
        strAdded$ = strAdded$ & str$ & ", "
      End If
      rstReqFlds.MoveNext
    Next r
  End If
  
  r = r - 1
  str$ = ""
  rstReqFlds.Close
  rstCoFlds.Close
  dbs.TableDefs.Refresh
  dbs.Close
  
  If Len(strAdded$) > 0 Then
    strAdded$ = Left(strAdded$, Len(strAdded$) - 2) 'remove trailing comma + space
    strName$ = DLookup("[txtName]", "tblCompanies", "[lngCompID]=" & lngCurrentComp)
    str$ = strName$ & vbCrLf & "COMPANY FIELDS VALIDATION:" & vbCrLf & vbCrLf
    If r = 1 Then
      str$ = str$ & "This field was missing and was added to the Company " & _
                    "Fields Table because it is required: " & vbCrLf & vbCrLf
    Else
      str$ = str$ & "These fields were missing and were added to the Company " & _
                    "Fields Table because they are required: " & vbCrLf & vbCrLf
    End If
    str$ = str$ & strAdded$
    DoCmd.OpenForm cstFRM_Msg, , , , , , str$
  End If
          
ValidateFieldsExit:
  Set rstReqFlds = Nothing
  Set rstCoFlds = Nothing
  Set dbs = Nothing
  Exit Sub
  
ValidateFieldsErr:
  If Err = 91 Then 'obj not set
    Resume Next
  Else
    MsgBox "Error: " & Err & ", " & Err.Description, vbExclamation, cstMDL & ":ValidateFields"
  End If
  Resume ValidateFieldsExit
  
End Sub

Function GetCoFieldsArray(lngCoID As Long, strQryName As String, varOut As Variant) As Integer
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim i As Integer
  i = 0
  On Error GoTo GetCoFieldsArrayErr
  
  If Not CompHasFields(lngCoID) Then Call SetNewCoFields(lngCoID)
  
  Set qdf = CurrentDb.QueryDefs(strQryName)
  qdf.Parameters(0) = lngCoID
  Set rst = qdf.OpenRecordset
  If rst.AbsolutePosition = -1 Then
    Debug.Print "GetCoFieldsArray (err): No records in : " & strQryName & " for " & lngCoID
    i = -1
    rst.Close
    qdf.Close
    GoTo GetCoFieldsArrayExit
  End If
  rst.MoveLast
  rst.MoveFirst

  ReDim varOut(i)
  Do While Not rst.EOF
    ReDim Preserve varOut(i)
    varOut(i) = rst(2) '= txtDispName
    i = i + 1
    rst.MoveNext
  Loop
  i = 0
  rst.Close
  qdf.Close
  
GetCoFieldsArrayExit:
  GetCoFieldsArray = i
  Set qdf = Nothing
  Set rst = Nothing
  Exit Function

GetCoFieldsArrayErr:
  i = Err.Number
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , cstMDL & ": GetCoFieldsArray"
  Resume GetCoFieldsArrayExit
End Function

Function SetNewCoFields(lngCoID As Long)
' Pre: lngCoID does NOT exist in tblCompFields
' Post: all the required fields for that company are created
'
  Dim dbs As DAO.Database
  Dim qdf As DAO.QueryDef
  Dim rstReqFlds As DAO.Recordset, rstCoFlds As DAO.Recordset
  Const cstReqFldsQry = "qryRequiredFields"
   
  On Error GoTo SetNewCoFieldsErr
  If CompHasFields(lngCoID) Then Exit Function
  
  Set dbs = CurrentDb
  Set qdf = dbs.QueryDefs(cstReqFldsQry)
  Set rstReqFlds = qdf.OpenRecordset
  rstReqFlds.MoveLast
  rstReqFlds.MoveFirst
  
  Set rstCoFlds = dbs.TableDefs("tblCompFields").OpenRecordset
  rstCoFlds.MoveLast

  Do While Not rstReqFlds.EOF
    rstCoFlds.AddNew
    rstCoFlds!lngCompID = lngCoID
    rstCoFlds!lngFldID = rstReqFlds!lngFldID
    rstCoFlds!lngAcctgTypeID = rstReqFlds!lngAcctgTypeID
    If Not IsNull(rstReqFlds!lngPriorityID) Then rstCoFlds!lngPriorityID = rstReqFlds!lngPriorityID
    rstCoFlds.Update
    rstReqFlds.MoveNext
  Loop
  rstReqFlds.Close
  rstCoFlds.Close
  qdf.Close
  dbs.Close
  
SetNewCoFieldsExit:
  Set qdf = Nothing
  Set rstReqFlds = Nothing
  Set rstCoFlds = Nothing
  Set dbs = Nothing
  Exit Function
  
SetNewCoFieldsErr:
  MsgBox "Error: " & Err & ", " & Err.Description, vbExclamation, cstMDL & ":SetNewCoFields"
  Resume SetNewCoFieldsExit
End Function

Function UpdateDateFromCal(Optional blnUndo As Boolean)
  Dim frmCal As Form
  Dim frmUpd As Form
  Dim ctl As Control
  Dim strF As String, strC As String
  
  On Error GoTo DateFromCalErr
  Set frmCal = Forms(cstFRM_Cal)
  With frmCal
    If (Len(!txtCtl) = 0 Or Len(!txtForm) = 0) Then
      GoTo DateFromCalExit  'both are needed
    End If
    strC = !txtCtl
    strF = !txtForm
    If Left(strF, 3) = "sfr" Then 'subform object
      Set frmUpd = Forms(cstFRM_Main)!sfrmAny.Form
    Else
      Set frmUpd = Forms(strF)
    End If
    
    Set ctl = frmUpd.Controls(strC)
    If IsMissing(blnUndo) Or Not blnUndo Then
       ctl = !acxCal.Value
    Else
      If blnUndo Then
        If CDate(!PreVal) <> Date Then
          ctl = Format(!PreVal, "Medium Date")
        Else
          ctl.Value = ctl.OldValue
        End If
      End If
    End If
  End With
  
DateFromCalExit:
  Set ctl = Nothing
  Set frmUpd = Nothing
  Set frmCal = Nothing
  Exit Function

DateFromCalErr:
  MsgBox "Error: (" & Err & ") " & Err.Description, vbExclamation, cstMDL & ": UpdateDateFromCal"
  Resume DateFromCalExit
End Function

Function GetMonthEndDate(varDateInput As Variant) As Date
  If Not IsDate(varDateInput) Then  'use current date
    GetMonthEndDate = DateSerial(Year(Date), Month(Date) + 1, 0)
  Else
    GetMonthEndDate = DateSerial(Year(varDateInput), Month(varDateInput) + 1, 0)
  End If
End Function

Function GetCoDefaultBudgetDesc(lngCo As Long) As String
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim str As String
  
  str = "<<no description>>"
  Set qdf = CurrentDb.QueryDefs("qryCoDefaultFRCDesc")
  qdf.Parameters(0) = lngCo
  Set rst = qdf.OpenRecordset
  If rst.AbsolutePosition <> -1 Then
    rst.MoveFirst
    str = rst(0)
  End If
  GetCoDefaultBudgetDesc = str
    
  rst.Close
  qdf.Close
  Set rst = Nothing
  Set qdf = Nothing
End Function

Function CoHasDefaultBudget(lngCo As Long) As Boolean
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  CoHasDefaultBudget = False
  
  Set qdf = CurrentDb.QueryDefs("qryCoDefaultBudget")
  qdf.Parameters(0) = lngCo
  Set rst = qdf.OpenRecordset
  If rst.AbsolutePosition <> -1 Then
    rst.MoveFirst
    If Not IsNull(rst(0)) Then CoHasDefaultBudget = True
  End If
  rst.Close
  qdf.Close
  Set rst = Nothing
  Set qdf = Nothing

End Function

Public Sub OpenCalendar(frmCaller As Form, strCallCtl As String)
  Dim blnBadCtl As Boolean
  Dim dteOpenDate  As Date
  Dim frmCal As Form
  Dim txt As TextBox
  Dim lngMoveRight As Long
  Dim str$
  Const cstMoveDown = 2200
  
  blnBadCtl = False
  On Error GoTo OpenCalErr
  
  ' Check if calling ctl valid by using naming convention:
  If InStr(1, strCallCtl, "dte", vbTextCompare) = 0 Then
    blnBadCtl = True
    GoTo OpenCalExit
  End If
  
  str$ = frmCaller.Name
  Set txt = frmCaller.Controls(strCallCtl)
  If IsNull(txt.Value) Or Not IsDate(txt.Value) Then
    dteOpenDate = Date
  Else
    dteOpenDate = CDate(txt.Value)
  End If

  DoCmd.OpenForm cstFRM_Cal, , , , , acHidden, dteOpenDate
  Set frmCal = Forms(cstFRM_Cal)
  lngMoveRight = (Screen.ActiveForm.WindowWidth / 2) - (frmCal.WindowWidth / 2)
  DoCmd.MoveSize Right:=lngMoveRight, down:=cstMoveDown

  frmCal.Visible = True
  With frmCal
    !acxCal = dteOpenDate
    !txtForm = str$
    !txtCtl = strCallCtl
  End With
  
OpenCalExit:
  Set txt = Nothing
  Set frmCal = Nothing
  Set frmCaller = Nothing
  
  If blnBadCtl Then
    MsgBox "Error: the calling control '" & strCallCtl & "' is not a date field (or named like one).", _
           vbExclamation, cstMDL & ":OpenCalendar"
  End If
  Exit Sub

OpenCalErr:
  MsgBox "Error: (" & Err & ") " & Err.Description & vbCrLf & _
         "Calling Form: " & str$ & vbCrLf & _
         "Calling Control: " & strCallCtl, vbExclamation, cstMDL & ":OpenCalendar"
  Resume OpenCalExit
End Sub

Sub MainFormSelForecast_AfterUpdate()
  Dim strSubformName As String, strCo As String
  Dim iSync As Integer
  strSubformName = ""
  '----------------------------------------------------------
  With Forms(cstFRM_Main)
    !lblDefBud.Visible = False: !txtDefaultFRCDesc = ""
    strCo = !cbxSelComp.Column(1)
    strSubformName = !sfrmAny.Form.Name
    !cbxSelForecast.Requery
    blnNoFrc = (!cbxSelForecast.ListCount = 0)
   
    If blnNoFrc Then
      MsgBox cstNoFrcForThisComp & ": " & vbCrLf & _
              strCo, vbExclamation, "No Forecasts Available"
      lngCurrentComp = !cbxSelComp
      !cbxSelForecast = ""
      lngCurrentForecast = 0
    Else
      ' Check if no frc selection, use first in list:
      !cbxSelForecast.SetFocus
      If (Len(!cbxSelForecast.Text) = 0) Then !cbxSelForecast = !cbxSelForecast.ItemData(0)
      lngCurrentForecast = !cbxSelForecast
      !lblDefBud.Visible = Nz(!cbxSelForecast.Column(3), 0)
      !txtDefaultFRCDesc = GetCoDefaultBudgetDesc(lngCurrentComp)
    End If

    Select Case strSubformName
        
      Case cstSFRM_Folio, cstSFRM_Fore
        If blnNoFrc Then
          iSync = SyncSubRecord("lngCompID", lngPrevCo)
        Else
          iSync = SyncSubRecord("lngCompID", lngCurrentComp, _
                              "lngForecastID", lngCurrentForecast) 'move pointer
          If iSync = 1 Then iSync = SyncSubRecord("lngCompID", lngPrevCo)
        End If
        
      Case cstSFRM_Summary
        With !sfrmAny.Form
          !cbxSelPeriodEnd.Requery
          !cbxSelPeriodEnd = !cbxSelPeriodEnd.ItemData(0)
          !lblTitlePeriod.Caption = !cbxSelPeriodEnd.Column(1)
          !txtCurrentCompany = strCo
          If blnNoFrc Then
            !txtSmryTitle = "<<No Forecast>>"
          Else
            !txtSmryTitle = .Parent!cbxSelForecast.Column(1) & _
                            " vs. " & .Parent!txtDefaultFRCDesc
          End If
          !ocxFlxGridForm.Rows = 1
          !ocxFlxGridForm.Clear
          !ocxFlxGridForm.ColWidth(0) = cstNamesColW
        End With
        
        If Not blnNoFrc And !cbxSelForecast.Column(3) <> 0 Then
          MsgBox "You have selected the default budget forecast" & vbCrLf & _
                 "...to be compared with itself.", vbExclamation, "Please select another forecast."
          !cbxSelForecast.SetFocus
        End If
        Call ResetSubForm(cstSFRM_Summary)
        
      Case cstSFRM_Series
        !sfrmAny.SourceObject = cstSFRM_Series  'reset subform
      
      Case cstSFRM_Analysis
        Call ResetAnalysisPage

    End Select
  End With
End Sub

Sub MainFormSelComp_AfterUpdate()
  Dim strSubformName As String, str As String
  Dim dteFrom As Date, dteTo As Date
  Dim iSync As Integer
  strSubformName = ""
  
  If blnAddForecast Then Exit Sub
  
  With Forms(cstFRM_Main)
    !cbxSelComp.SetFocus
    !cbxSelComp.SelLength = 0
    lngCurrentComp = !cbxSelComp
    If lngPrevCo = lngCurrentComp Then Exit Sub
    
    strSubformName = !sfrmAny.Form.Name
    !cbxSelForecast = ""
    !cbxSelForecast.Requery
    blnNoFrc = (!cbxSelForecast.ListCount = 0)
    !cbxSelForecast.SetFocus
    ' Select first in list if none selected:
    If (Not blnNoFrc And Len(!cbxSelForecast.Text) = 0) Then !cbxSelForecast = !cbxSelForecast.ItemData(0)
    !cbxSelForecast.SelLength = 0
    !txtDefaultFRCDesc = GetCoDefaultBudgetDesc(lngCurrentComp)

    Select Case strSubformName
      Case cstSFRM_Folio, cstSFRM_Fore, cstSFRM_Series, cstSFRM_Summary
        Call MainFormSelForecast_AfterUpdate
        
      Case cstSFRM_Analysis
        Call ResetAnalysisPage
   
      Case cstSFRM_Comp
        iSync = SyncSubRecord("lngCompID", lngCurrentComp) 'move pointer
    End Select
  End With
End Sub

Sub OpenExportForm()
  If IsNull(Forms(cstFRM_Main)!cbxSelForecast) Then
    MsgBox Forms(cstFRM_Main)!cbxSelCompID.Column(1) & _
      " does not have any forecasts!", vbExclamation, "No Forecasts to Export"
  Else
    DoCmd.OpenForm "frmExport"
  End If
End Sub

Function OpenClosingInfo() As Integer
  Dim iCancel As Integer
  iCancel = 0: blnNewClosing = False

  If lngCurrentComp = 0 Then lngCurrentComp = Forms(cstFRM_Main)!cbxSelComp.Column(0)
  If IsNull(DLookup("[lngCompID]", "tblClosingInfo", "[lngCompID]=" & lngCurrentComp)) Then
    If MsgBox("There is no closing information for " & Forms(cstFRM_Main)!cbxSelComp.Column(1) & ". " & vbCrLf & _
              "Do you want to enter it?", vbExclamation + vbYesNo, "No Closing Info") = vbYes Then
      blnNewClosing = True
    Else
      iCancel = 1
      Exit Function
    End If
  End If
  DoCmd.Close acForm, cstFRM_Closing
  OpenClosingInfo = iCancel
  DoCmd.OpenForm cstFRM_Closing, , , , , , blnNewClosing
End Function



