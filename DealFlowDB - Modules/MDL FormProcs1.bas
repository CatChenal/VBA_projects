Attribute VB_Name = "MDL FormProcs1"
Option Compare Database
Option Explicit
'===============================================================================
' DBFrontEnd April-7-03 15:20
' MDL FormProcs
' Prev: Nov-22-02 16:45
'===============================================================================
Public blnRequeryCbx As Boolean
Public Const cstMillionMultiple = 1000000
Const cstMDL = "MDL FormsProcs"

Function ResetToolbars()
  DoCmd.Maximize
  DoCmd.ShowToolbar "Database", acToolbarNo
  DoCmd.ShowToolbar "Formatting (Form/Report)", acToolbarNo
  DoCmd.ShowToolbar "Formatting (Datasheet)", acToolbarNo
  DoCmd.ShowToolbar "Form View", acToolbarNo
End Function

Function CheckFormArgs(frm As Form) As String
  If Len(strCallingForm) = 0 Then
    If Len(frm.OpenArgs & "") = 0 Then
      strCallingForm = cstMainForm
    Else
      strCallingForm = frm.OpenArgs
    End If
  End If
  CheckFormArgs = strCallingForm
End Function

Sub CloseFrm(frmCallingForm As Form)
  Dim frm As Form
  Dim strFrm As String
  
  On Error GoTo CloseFrmErr
  strFrm = frmCallingForm.Name
  If strCallingForm = "" Then strCallingForm = CheckFormArgs(frmCallingForm)

  If frmCallingForm.Dirty Then Call SaveRec(frmCallingForm)
  
  If strFrm = cstDealForm Then
    If IsLoaded(cstIssuerForm) Then Forms(cstIssuerForm).Visible = False
    If IsLoaded(cstFinStatusForm) Then Forms(cstFinStatusForm).Visible = False
    If IsLoaded(cstSourceForm) Then DoCmd.Close acForm, cstSourceForm
  End If
       
  If blnNewSource Then
    If strCallingForm = cstMainForm Then Forms(cstMainForm).Refresh
  End If
  
  frmCallingForm.Visible = False
  DoCmd.OpenForm strCallingForm
  If strFrm = cstSourceForm Then DoCmd.Close acForm, strFrm
 
CloseFrmExit:
  Set frm = Nothing
  Set frmCallingForm = Nothing
  Exit Sub

CloseFrmErr:
  strCallingForm = cstMainForm
 If Err.Number <> 2455 Then
    MsgBox "Error: (" & Err & ") " & Err.Description & vbCrLf & _
           "Form: " & strFrm, vbExclamation, cstMDL & " CloseFrm"
    Err.Clear
  End If
  Resume CloseFrmExit
End Sub

Sub SaveRec(frmCallingForm As Form)
  On Error GoTo SaveRecErr
  
  strCallingForm = CheckFormArgs(frmCallingForm)
  If frmCallingForm.Dirty Then
    If MsgBox(cstSaveMsg, vbQuestion + vbYesNo, "Save?") = vbYes Then
      frmCallingForm.SetFocus
      DoCmd.RunCommand acCmdSaveRecord
      blnRequeryCbx = True
      If Nz(frmCallingForm.OpenArgs, cstMainForm) = cstMainForm Then Forms(cstMainForm).Requery
      If frmCallingForm.Modal Then frmCallingForm.Modal = False
    Else
      blnRequeryCbx = False
      'frmCallingForm.Controls(0).SetFocus
      If frmCallingForm.NewRecord Then Call UndoRec(frmCallingForm)
    End If
  End If
  
SaveRecExit:
  Set frmCallingForm = Nothing
  Exit Sub
  
SaveRecErr:
  If Err.Number <> 2455 Then
    MsgBox "Error: (" & Err & ") " & Err.Description & vbCrLf & _
           "Form: " & frmCallingForm.Name, vbExclamation, cstMDL & ": SaveRec"
    Err.Clear
  End If
  Resume SaveRecExit
End Sub

Sub UndoRec(frmCallingForm As Form)
  Dim ctl As Control
  Dim strForm As String
  
  On Error GoTo UndoRecErr
  
  strForm = frmCallingForm.Name
  
  If frmCallingForm.NewRecord = True Then
    If MsgBox("Do you want do undo this new record?", _
               vbQuestion + vbYesNo, "Undo New?") = vbYes Then
      DoCmd.RunCommand acCmdUndo
      If frmCallingForm.RecordsetClone.RecordCount > 1 Then
        DoCmd.GoToRecord acDataForm, strForm, acPrevious
      Else
        DoCmd.GoToRecord acDataForm, strForm, acFirst
      End If
    End If
  Else
    If frmCallingForm.Dirty Then
      On Error Resume Next
      If MsgBox("Do you want do undo your unsaved changes?", vbQuestion + vbYesNo, "Undo changes?") = vbYes Then
        For Each ctl In frmCallingForm.Controls
          If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then  ' Restore
            If ctl.Name <> "txtContact" And ctl.Name <> "txtContact2" Then ctl.Value = ctl.OldValue
            If Err <> 0 Then
              If Err = 2448 Then 'oldvalue doesn't apply
                Err.Clear
                Resume Next
              Else
                GoTo UndoRecErr
              End If
            End If
          End If
        Next ctl
        frmCallingForm.SetFocus
      End If
    Else
      If MsgBox("No unsaved change to undo: do you want to delete this record instead?", _
                vbQuestion + vbYesNo, "Undo Record") = vbYes Then
        DoCmd.RunCommand acCmdDeleteRecord
        Forms(cstMainForm).Requery
      End If
   End If
  End If
  
UndoRecExit:
  Set ctl = Nothing
  Set frmCallingForm = Nothing
  Exit Sub

UndoRecErr:
  If Err.Number <> 2455 Then
    If Err <> 2046 Then 'not available: nothing to undo
      MsgBox "Error: (" & Err & ") " & Err.Description & vbCrLf & _
             "Form: " & frmCallingForm.Name, vbExclamation, cstMDL & " UndoRec"
    Else
      MsgBox "No change to undo." & vbCrLf & _
             "Form: " & frmCallingForm.Name, vbInformation, cstMDL & " UndoRec"
    End If
    Err.Clear
  End If
  Resume UndoRecExit
End Sub

Sub NewRec(frmCallingForm As Form)
  Dim strForm As String
  On Error GoTo NewRecErr

  strForm = frmCallingForm.Name
  If frmCallingForm.Dirty Then
    If MsgBox("Do you want to save this record before adding a new one?", _
               vbQuestion + vbYesNo, "Save Current & Add?") = vbYes Then
      Call SaveRec(frmCallingForm)
    Else
      Call UndoRec(frmCallingForm)
    End If
  End If
  DoCmd.GoToRecord acDataForm, strForm, acNewRec
  If strForm = cstSourceForm Then blnNewSource = True
  If strForm = cstDealForm Then blnNewDeal = True
  If strForm = cstIssuerForm Then blnNewIssuer = True
  
NewRecExit:
  Set frmCallingForm = Nothing
  Exit Sub

NewRecErr:
  MsgBox "Error: (" & Err & ") " & Err.Description & vbCrLf & _
         "Form: " & frmCallingForm.Name, vbExclamation, cstMDL & " NewRec"
  Err.Clear
  Resume NewRecExit
End Sub

Sub ShowStatusForm(strParentForm As String, Optional blnAdd As Boolean)
  Dim frmParent As Form
  Dim strFrom As String, strCrit As String, strMsg As String
  Dim bln As Boolean
  Dim var As Variant
  Dim lngDeal As Long, lngIssuer As Long
  Dim dte As Date
  
  On Error GoTo ShowStatusFormErr

  bln = FinRecExists
  If (strParentForm <> cstDealForm) Then Exit Sub
  If Not IsLoaded(strParentForm) Then Exit Sub
  
  ' Get params:
  Set frmParent = Forms(strParentForm)
  lngDeal = frmParent!lngDealNum
  lngIssuer = frmParent!lngDealIssuerNum
  dte = GetPrevQtrEndDate(frmParent!dteDealDateIn)
  strCrit = "[lngFinIssuerNum] = " & lngIssuer & " AND [lngFinDealNum] = " & lngDeal
  '
  If Not bln Then
      'add rec if OK:
    strMsg = "This Issuer does not have any Financials record for this Deal." & vbCrLf
    strMsg = strMsg & "Do you want to create one?"
    If MsgBox(strMsg, vbQuestion + vbYesNo, "No Issuer Financials for Current Deal") = vbYes Then
      Call AddFinStatusRec(strParentForm, lngIssuer, lngDeal, dte)
    End If
  Else
    DoCmd.OpenForm cstFinStatusForm, , , strCrit, , , cstDealForm 'strParentForm
  End If
  
ShowStatusFormExit:
  Set frmParent = Nothing
  Exit Sub
  
ShowStatusFormErr:
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstMDL & ": ShowStatusForm"
  Resume ShowStatusFormExit
End Sub

Function FinRecExists() As Boolean
' Form frmDeal has to be open
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim frm As Form
  
  On Error GoTo FinRecExistsErr
  FinRecExists = False
  If IsLoaded(cstDealForm) Then
    Set frm = Forms(cstDealForm)
    If IsNull(frm!lngDealIssuerNum) Or IsNull(frm!lngDealNum) Then Exit Function
  Else
    MsgBox "Cannot check: the Deal Entry Form is not open.", vbExclamation, _
            "Check for Issuer's Financial Data"
    Exit Function
  End If
    
  If Not IsSet(dbs) Then Set dbs = CurrentDb
  Set qdf = dbs.QueryDefs("qryIssuerFinStatus")
  qdf.Parameters(0) = frm!lngDealIssuerNum
  qdf.Parameters(1) = frm!lngDealNum
  Set rst = qdf.OpenRecordset
  If rst.AbsolutePosition <> -1 Then
    rst.MoveLast
    If rst.RecordCount > 0 Then FinRecExists = True
  End If
  
FinRecExistsExit:
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Exit Function
  
FinRecExistsErr:
  FinRecExists = False
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstMDL & ": FinRecExists"
  Resume FinRecExistsExit
End Function

Sub AddFinStatusRec(strParent As String, lngIssuer As Long, lngDeal As Long, dte As Date)
  Dim frmStatus As Form
  
  strCallingForm = strParent
  If Not IsLoaded(cstFinStatusForm) Then
    DoCmd.OpenForm cstFinStatusForm, , , , acFormAdd, , strParent
  End If
  Set frmStatus = Forms(cstFinStatusForm)
  With frmStatus
    .Visible = True
    !lngFinIssuerNum = lngIssuer
    !lngFinDealNum = lngDeal
    !dteFinPeriodEnd = dte
    'set all other ctl to null
  End With
  frmStatus.Visible = True
  Set frmStatus = Nothing
End Sub

Public Function DealSourceNameUpdate(cbxUsed As ComboBox) As String
' For frmDeal only: datasource for txtContact (Source) & txtContact2 (Eq.Sponsor)
' No check on txt ctl receiving it.
'
  Dim str As String
  On Error GoTo DealSourceNameUpdateErr
 
  If Not IsNull(cbxUsed) Then
  ' col(0): source num; col(1):sourcename;  col(2): last; col(3): first.
    If IsNull(cbxUsed.Column(3)) Then  ' no first name
      If IsNull(cbxUsed.Column(2)) Then ' no last name
        str = "<no record>" '"<no contact's name>"
      Else
       str = cbxUsed.Column(2) 'last name only: OK
      End If
    Else
      If IsNull(cbxUsed.Column(2)) Or _
                cbxUsed.Column(2) = "" Then ' first name only
        str = "<no last name>, " & cbxUsed.Column(3)
      Else 'both names
        str = cbxUsed.Column(2) & ", " & cbxUsed.Column(3)
      End If
    End If
  Else
   str = "<no record>"
  End If
  DealSourceNameUpdate = str
  
DealSourceNameUpdateExit:
  Exit Function
  
DealSourceNameUpdateErr:
  Set cbxUsed = Nothing
  DealSourceNameUpdate = "<>"
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstMDL & ": DealSourceNameUpdate"
  Resume DealSourceNameUpdateExit
End Function

Public Function SourceNotInList(strNewData As String, strOrigin As String) As Integer
'For the Source or Equity Sponsor comboxes on frmDeal only.
'Returns the Response expected by the NotInList event (if LimitToList=Yes).
' strOrigin refer to either 'Source' or 'Equity Sponsor'.
  Dim rst As DAO.Recordset
  Dim frmNewSource As Form
  Dim strMsg As String, strTitle As String
  Dim strRespLast As String, strRespFirst As String
  Dim blnProceed As Boolean
  
  On Error GoTo SourceNotInListErr
  SourceNotInList = 0
  blnProceed = False
  
  If ((strOrigin = "Equity Sponsor") Or (strOrigin = "Source")) Then
    blnProceed = True
  Else
    MsgBox "Error: " & vbCrLf & _
           " Erroneous call of 'SourceNotInList' function, or" & vbCrLf & _
           " Either one of the combo boxes for Source and Equity Sponsor" & vbCrLf & _
           " on the Deal Details form (frmDeal) has been renamed.", vbExclamation, _
           cstMDL & ": SourceNotInList-bad call"
    Exit Function
  End If
  
  strMsg = strNewData & " is not on file. " & vbCrLf & "Is this a new " & strOrigin & "?"
  strTitle = "Unknown " & strOrigin
  If MsgBox(strMsg, vbYesNo + vbQuestion, strTitle) = vbYes Then
    
InputStart:
    strRespLast = Trim(InputBox("Please, enter the contact's last name:", _
                           "New Contact", "<last name>"))
    If strRespLast = "" Then
      If MsgBox("The last name is required." & vbCrLf & _
                 "Do you want to cancel this new source entry?", vbExclamation + vbYesNo, _
                 "Missing Last Name") = vbNo Then
        blnProceed = True
        GoTo InputStart
      Else
        blnProceed = False
      End If
    End If
      
    If Not blnProceed Then Exit Function
    strRespFirst = Trim(InputBox("Please, enter the contact's first name:", _
                                 "New Contact", "<first name>"))
    
    'Ref form w/o showing it:
    If Not IsLoaded(cstSourceForm) Then
      DoCmd.OpenForm cstSourceForm, acNormal, , , , , cstDealForm
    End If
    Set frmNewSource = Forms!frmSource
    ' Add new rec in rst:
    Set rst = frmNewSource.RecordsetClone
    blnNewSource = True
    With rst
      .AddNew
      !txtSourceName = strNewData
      !txtSourceContactLast = strRespLast
      !txtSourceContactFirst = strRespFirst
      !lngSourceTypeNum = 9 '  sponsor group
      .Update
      .Bookmark = .LastModified
    End With
    frmNewSource.Bookmark = rst.Bookmark
    SourceNotInList = 2 'acDataErrAdded
  Else
    SourceNotInList = 0 ' acDataErrContinue
    blnNewSource = False
  End If
    
  If blnNewSource Then
    If strOrigin = "Source" Then
      Forms(cstDealForm)!cbxDealSourceName = rst!lngSourceNum
    Else
      Forms(cstDealForm)!cbxEqSponsor = rst!lngSourceNum
    End If
  End If
  
SourceNotInListExit:
  Set frmNewSource = Nothing
  Set rst = Nothing
  Exit Function
  
SourceNotInListErr:
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstMDL & ": SourceNotInList"
  Resume SourceNotInListExit
End Function

Function DeleteCurrentRec(ByVal lngDealID As Long)
'Pre: lngDealID is the unique id of a deal.
  Dim rst As DAO.Recordset
  Dim tbl As DAO.TableDef
  Dim strSQL As String, strMsg As String
      
  On Error GoTo DeleteCurrentRecErr
  
  strMsg = "The deletion operation you requested will delete ALL records related to the " & _
           "current deal." & vbCrLf & "Proceed with deletion?"
  If MsgBox(strMsg, vbYesNo + vbQuestion, "Deletion confirmation") = vbNo Then Exit Function
  
  If IsLoaded(cstDealForm) Then DoCmd.Close acForm, cstDealForm

  Call DeleteRecordData(CurrentDb, lngDealID, "tblDeal")
  Forms!frmDealSelection.Refresh
  Forms!frmDealSelection.Requery
  
DeleteCurrentRecExit:
  Exit Function
DeleteCurrentRecErr:
  MsgBox Err.Number & Err.Description, , cstMDL & ": DeleteCurrentRec(" & lngDealID & ")"
  Resume DeleteCurrentRecExit
End Function

Function DeleteRecordData(db As DAO.Database, ByVal lngId As Long, strTable As String)
'Pre: If strTble=tblFunAllocation then lngID=lngSecDealNum, else it is the Deal ID.
  Dim rst As DAO.Recordset, rstSec As DAO.Recordset
  Dim strSQL As String, strSQLSec As String, strMsg As String
  Dim lngSec As Long
  Const cst = "SELECT * FROM "
  
  On Error GoTo DeleteRecordDataErr
 
  Select Case strTable
    Case "tblSecurity"
       
      strSQLSec = cst & "tblSecurity WHERE tblSecurity.lngSecDealNum = " & lngId & ";"
      Set rstSec = db.OpenRecordset(strSQLSec)
      With rstSec
        If .AbsolutePosition <> -1 Then
          .MoveFirst
          Do While Not .EOF
            lngSec = !lngSecDealNum
            Call DeleteRecordData(db, lngSec, "tblFundAllocation")
            .Delete
            .MoveNext
          Loop
        End If
      End With
     'strSQL = cst & "tblSecurity WHERE tblSecurity.lngSecDealNum = " & lngID & ";"
      
    Case "tblFundAllocation"
      strSQL = cst & "tblFundAllocation WHERE tblFundAllocation.lngAllocSecNum = " & lngId & ";"

    Case "tblFinStat"
      strSQL = cst & "tblFinStat WHERE tblFinStat.lngFinDealNum = " & lngId & ";"
    
    Case "tblDeal"
      Call DeleteRecordData(db, lngId, "tblFinStat")
      Call DeleteRecordData(db, lngId, "tblSecurity")
      strSQL = cst & "tblDeal WHERE tblDeal.lngDealNum = " & lngId & ";"
      
  End Select
  
  If strTable = "tblSecurity" Then Exit Function 'recs already deleted at this point
  
  Set rst = db.OpenRecordset(strSQL, dbOpenDynaset)
  With rst
    If .AbsolutePosition <> -1 Then
      .MoveFirst
      Do While Not .EOF
        .Delete
        .MoveNext
      Loop
    End If
  End With
  
DeleteRecordDataExit:
  Set rst = Nothing
  Set rstSec = Nothing
  Exit Function
  
DeleteRecordDataErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , cstMDL & ": DeleteRecordData(" & lngId & ")"
 Resume DeleteRecordDataExit
End Function

Function CheckAmtEntry(ctlAmt As Control, lngAmtMultiple As Long) As Single
' To check whether the amount was entered according to 2 possible 'format',
' either as $thousands, or $millions multiple.
'
  Dim strMsg As String, strMultiple As String, strTitle As String, strName As String
  Dim sglVal As Single
  Dim iNumberOfZeros As Integer
  strMsg = "": sglVal = 0
  Debug.Print "MDL FormProcs CheckAmtEntry"
  
  If lngAmtMultiple < 1000 Then Exit Function
  
  If Not IsNull(ctlAmt) And (ctlAmt <> 0) Then
    sglVal = ctlAmt.Value
    
    Select Case lngAmtMultiple
      Case 1000
        strMultiple = "thousands"
        iNumberOfZeros = 3
      Case 1000000
        strMultiple = "millions"
        iNumberOfZeros = 6
    End Select
  
    If sglVal * lngAmtMultiple >= 10 ^ (2 * iNumberOfZeros) Then
                'e.g: if strMultiple = "thousands":
                'sglVal = 2,000; Multiple = 1,000
                'sglVal * Multiple = 2,000,000 > millions
                '
      ' Has it been entered in disregard of the amount entry convention?
      strName = Mid$(ctlAmt.Name, 4) 'strName = ctlAmt.Tag
      'Debug.Print "ctlAmt.Name: " & ctlAmt.Name, "ctlAmt.Tag: " & ctlAmt.Tag
      
      strTitle = StrConv(strMultiple, vbProperCase) & " Multiple Check"
      strMsg = "Amounts are stored in " & strMultiple & " of dollars." & vbCrLf & vbCrLf
      strMsg = strMsg & "The amount in " & strName & " is " & Format(sglVal, "$#,##0.0")
      strMsg = strMsg & "." & vbCrLf & "Do you want to adjust it accordingly?"
      
      If MsgBox(strMsg, vbExclamation + vbYesNo, strTitle) = vbYes Then
        Do While sglVal >= lngAmtMultiple
          sglVal = sglVal / lngAmtMultiple
        Loop
      Else
        sglVal = 0
      End If
    End If
  End If
  CheckAmtEntry = sglVal
End Function

Function BillionAmtCheck(ctlAmt As Control, lngAmtMultiple As Long) As Single
' To check whether an amount greater than 1 billion was entered according to
' two possible conventions: either as $thousands, or $millions multiple.
'
  Dim strMsg As String, strMultiple As String, strName As String
  Dim sglVal As Single, sglActualVal As Single
  strMsg = "": sglVal = 0
  
  If lngAmtMultiple < 1000 Then Exit Function
  
  If Not IsNull(ctlAmt) And (ctlAmt <> 0) Then
    sglVal = ctlAmt.Value
    sglActualVal = sglVal * lngAmtMultiple   ' calculate the amount as per convention
    If sglActualVal >= 1000000000 Then       ' over 1 billion Then
      ' Has it been entered in error?
      Select Case lngAmtMultiple
        Case 1000
          strMultiple = "thousands"
        Case 1000000
          strMultiple = "millions"
      End Select
      strName = Mid$(ctlAmt.Name, 4)
      
      strMsg = "By convention in this database, currency amounts are stored in " & _
               strMultiple & " of dollars." & vbCrLf & vbCrLf
      strMsg = strMsg & "The amount in " & strName & " is actually " & _
               Format(sglActualVal, "$#,###.0")
      strMsg = strMsg & " (" & Format(sglActualVal / 10 ^ 9, "$#,###.### billions") & _
               "." & vbCrLf & "Is this correct?"
      
      If MsgBox(strMsg, vbExclamation + vbYesNo, "Billion Amount Check") = vbNo Then
        sglVal = sglActualVal / 10 ^ 9
        'Debug.Print "Corrected val: " & sglVal, "ie: " & Format(sglVal, "$#,###.0")
        sglVal = CCur(Nz(InputBox("Either accept the proposed correction, " & _
                                 "or edit it and click OK.", _
                                 "Billion Amount Correction", sglVal), 0))
      End If
    End If
  End If
  BillionAmtCheck = sglVal
End Function

Function ClearCtlWithRightMouseDown(ctlActive As Control, iButton As Integer)
  If ctlActive.ControlType = acComboBox Or ctlActive.ControlType = acTextBox Then
    If iButton = acRightButton Then ctlActive = Null
  End If
End Function

