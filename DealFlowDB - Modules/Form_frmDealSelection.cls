VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDealSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'===============================================================================
' DBFrontEnd
' CLSMDL frmDealSelection Sep-05-03 14:40
'
'===============================================================================
Const cstThisForm = cstMainForm 'for err msg

Private Sub cmbNewDeal_Click()
  Dim frm As Form
  On Error GoTo cmbNewDealErr
  
  strCallingForm = cstThisForm
  If Not IsLoaded(cstDealForm) Then DoCmd.OpenForm cstDealForm, , , , , , strCallingForm
  Forms(cstDealForm).Visible = True
  Set frm = Forms(cstDealForm)
  Call NewRec(frm)
  frm.SetFocus
  
cmbNewDealExit:
  Set frm = Nothing
  Exit Sub

cmbNewDealErr:
  MsgBox Err.Number & ": " & Err.Description, , cstThisForm & ": cmbNewDeal"
  Resume cmbNewDealExit
End Sub

Private Sub cmbDelete_Click()
  Call DeleteCurrentRec(Me!lngDealNum)
  Me.Requery
End Sub

Private Sub cmbToDealForm_Click()
  On Error GoTo cmbToDealForm_ClickErr
  
  If Not IsNull(Me!lngDealNum) Then
    strCallingForm = cstThisForm
    DoCmd.OpenForm cstDealForm, , , "[lngDealNum]=" & Me!lngDealNum, , , strCallingForm
  Else
    MsgBox "Move the pointer to a deal record prior to viewing its details.", _
           vbExclamation, "No current record"
  End If
  
cmbToDealForm_ClickExit:
  Exit Sub
  
cmbToDealForm_ClickErr:
  Resume cmbToDealForm_ClickExit
End Sub

Private Sub cmbToIssuerForm_Click()
  DoCmd.OpenForm cstIssuerForm, , , "[lngIssuerNum]=" & Me!lngDealIssuerNum, , , cstThisForm
End Sub

Private Sub cmbToSourceForm_Click()
  DoCmd.OpenForm cstSourceForm, , , "[lngSourceNum]= " & Me!lngDealSourceNum, , , cstThisForm
End Sub

Private Sub cmbToReportSelectionForm_Click()
  If Not IsLoaded(cstReportForm) Then DoCmd.OpenForm cstReportForm, , , , , acHidden
  Forms(cstReportForm).Visible = True
  Call ResetFilterClues(cstReportForm)
End Sub

Private Sub cmbExit_Click()
  If MsgBox("Close database?", vbQuestion + vbYesNo, "Exit...") = vbYes Then
    Set dbs = Nothing
    DoCmd.RunCommand acCmdAppMinimize
    DoCmd.RunCommand acCmdExit
  Else
    DoCmd.Hourglass True
    Call CloseAllOpenFormsButCaller(Me)
    DoCmd.Hourglass False
  End If
End Sub

Private Sub Form_Activate()
  DoCmd.Maximize
  Call ApplyFormFilter(cstThisForm)
End Sub

Private Sub FullIssuer_DblClick(Cancel As Integer)
  Call cmbToIssuerForm_Click
End Sub

Private Sub FullSource_DblClick(Cancel As Integer)
  Call cmbToSourceForm_Click
End Sub

Private Sub Form_Current()
  Dim strFilter As String
  Dim frm As Form
  
  On Error GoTo Form_CurrentErr
  blnFilterErr = False
  
  '' Refresh count
  Me!RecCount = Me.RecordsetClone.RecordCount
  Me!TotalRecs.Requery
  
  If IsLoaded(cstSourceForm) Then 'show source per current deal if visible
   Set frm = Forms(cstSourceForm)
   If frm.Visible = True Then  'synchronize
     If IsNull(Me!lngDealSourceNum) Then
        strFilter = "[lngSourceNum] = 9000"
     Else
        strFilter = "[lngSourceNum] = " & Me!lngDealSourceNum
     End If
     frm.Filter = strFilter
     frm.FilterOn = True
     Me.SetFocus
     Set frm = Nothing
   End If
  End If
  
  If IsLoaded(cstFinStatusForm) Then DoCmd.Close acForm, cstFinStatusForm
  
  If IsLoaded(cstIssuerForm) Then
    Set frm = Forms(cstIssuerForm)
    If frm.Visible = True Then
      If IsNull(Me!lngDealIssuerNum) Then
         strFilter = "[lngIssuerNum] = 9000"
      Else
         strFilter = "[lngIssuerNum] = " & Me!lngDealIssuerNum
      End If
      frm.Filter = strFilter
      frm.FilterOn = True
      Me.SetFocus
      Set frm = Nothing
    End If
  End If
   
Form_CurrentExit:
   Exit Sub
   
Form_CurrentErr:
  If Err.Number = 3021 Then   'no current rec
    Me!RecCount = 0
    blnFilterErr = True
  Else
   MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstThisForm & ": Form_Current"
   Resume Form_CurrentExit
  End If
End Sub

Private Sub lblFilter_Click()
  Call ShowFilterForm(cstMainForm, False)
End Sub

Private Sub txtSourceContactLast_DblClick(Cancel As Integer)
  Call cmbToSourceForm_Click
End Sub
