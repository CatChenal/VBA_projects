VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmUpdContDLG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
' DBFrontEnd
' CLSMDL frmUpdContDLG May-10-02 16:20
'
'================================================================================
Dim blnUpdOL As Boolean
Dim blnAddOL As Boolean
Const cstThisForm = "frmUpdContDLG"

Private Sub chkOLNew_AfterUpdate()
  blnAddOL = CBool(Me!chkOLNew.Value)
End Sub

Private Sub chkOLUpd_AfterUpdate()
  Me!cmbCancel.SetFocus
  
  blnUpdOL = CBool(Me!chkOLUpd.Value)
  If blnUpdOL Then
    Me!chkOLNew = (Me!chkOLUpd * 1)
  Else
    Me!cmbOK.Enabled = True
  End If
  Me!chkOLUpd.Visible = blnUpdOL
  Me!chkOLNew.Visible = blnUpdOL
End Sub

Private Sub cmbCancel_Click()
  DoCmd.Close acForm, cstThisForm
End Sub

Private Sub cmbUpdOK_Click()
  Dim strL As String, strF As String
  Dim frm As Form
  
  On Error GoTo Err_cmbOK_Click
  'MsgBox "strCallingForm : " & strCallingForm
  Set frm = Forms!frmManagement
  With frm
    strL = !txtMgrLast
    strF = !txtMgrFirst
  End With
  
  DoCmd.Hourglass True
  Me.Visible = False
  '
  Call SaveRec(frm)
  If blnUpdOL Then Call UpdateOLContact(frm, strL, strF, blnAddOL)
  '
  Set frm = Nothing
  DoCmd.Hourglass False
 
  DoCmd.Close acForm, cstThisForm
  
Exit_cmbOK_Click:
  Set frm = Nothing
  Exit Sub

Err_cmbOK_Click:
  DoCmd.Hourglass False
  strCallingForm = "frmContacts"
  MsgBox Err & ": " & Err.Description, , cstThisForm & " cmbOK"
  Resume Exit_cmbOK_Click
End Sub

Private Sub lblUpdate_Click()
  With Me
    !lblUpdOLCard.Visible = True
    !chkOLUpd.Visible = True
    !lblCreateOLCard.Visible = True
    !chkOLNew.Visible = True
    !cmbUpdOK.Enabled = True
  End With
End Sub
