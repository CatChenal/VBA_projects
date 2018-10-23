VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmIssuer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
' DBFrontEnd
' CLSMDL frmIssuer Jul-24-02
' Prev: May-10-02 12:15
' Changes: cleanup
'
'================================================================================
Const cstThisFrm = cstIssuerForm
Const cstThisFrmHeight = 1440 * 3.2    ' 1440 twips = 1 inch.
Const cstThisFrmHeightPlus = 1440 * 5.15  ' expanded
'
' NOTE: the Form.OpenArgs method is reserved for setting the global variable
'       strCallingForm, which is used by the form's close button so that the user
'       returns to the calling form upon closing this one.
'---------------------------------------------------------------------------------
Private Sub cmbClose_Click()
  Call SaveRec(Me)
  Me!txtIssuerName.SetFocus
  DoCmd.Close acForm, cstThisFrm, acSaveNo
End Sub

Private Sub cmbManagement_Click()
  Dim frmMgmt As Form
  Dim rst As DAO.Recordset
  Dim strCrit As String, strFirst As String, strLast As String, strMsg As String, strFull As String
  Dim lngCo As Long
  Dim intResponse As Integer
  
  On Error GoTo cmbManagementErr
  
  DoCmd.Hourglass True
  
  lngCo = Me!lngIssuerNum
  strCrit = "[lngCompanyID] = " & lngCo
  
  If Not IsLoaded(cstMgmtForm) Then
    DoCmd.OpenForm cstMgmtForm, , , , , acHidden
  End If
  Set frmMgmt = Forms(cstMgmtForm)
  frmMgmt.Visible = False
   
  Set rst = frmMgmt.RecordsetClone
  rst.FindFirst strCrit
  If Not rst.NoMatch Then
    frmMgmt.Bookmark = rst.Bookmark
    DoCmd.Hourglass False
    frmMgmt.SetFocus
    Me.Modal = False
    Me.Visible = False
    frmMgmt.Visible = True
  Else
    If IsNull(Me!txtIssuerContactLast) Then
      strFirst = "First": strLast = "Last"
      strMsg = "Do you want to enter a new Manager?" & vbCrLf
      intResponse = MsgBox(strMsg, vbQuestion + vbOKCancel, "No manager found")
    Else
      strLast = Me!txtIssuerContactLast
      strFull = strLast
      If IsNull(strFirst = Me!txtIssuerContactFirst) Then
        strFirst = ""
      Else
        strFirst = Me!txtIssuerContactFirst
        strFull = strFull & ", " & strFirst
      End If
      strMsg = "Do you want to enter " & strFull & " as a Manager?" & vbCrLf & _
               "Select 'No' to enter a different person;" & vbCrLf & _
               "       'Cancel' to prevent record creation."
      intResponse = MsgBox(strMsg, vbQuestion + vbYesNoCancel, "No manager found")
    End If
    
    If intResponse <> vbCancel Then
      If intResponse = vbNo Then  'reset name fields:
        strFirst = "First": strLast = "Last"
      End If
      'Add a mgmt record with current contact data
      With rst
        .AddNew
        !lngCompanyID = lngCo
        !txtMgrFirst = strFirst
        !txtMgrLast = strLast
        !txtMgrPhone = Me!txtIssuerContactPhone
        .Update
        frmMgmt.Bookmark = .LastModified
        DoCmd.Hourglass False
        frmMgmt.SetFocus
        Me.Modal = False
        Me.Visible = False
        frmMgmt.Visible = True
      End With
    Else
      DoCmd.Hourglass False
      DoCmd.Close acForm, cstMgmtForm
      GoTo cmbManagementExit
    End If
  End If
    
cmbManagementExit:
  Set frmMgmt = Nothing
  Set rst = Nothing
  Exit Sub
    
cmbManagementErr:
  DoCmd.Hourglass False
  MsgBox Err.Number & ": " & Err.Description, , cstThisFrm & ": cmbManagement"
  Resume cmbManagementExit
  
End Sub

Private Sub cmbNewIssuer_Click()
  RunCommand acCmdRecordsGoToNew
End Sub

Private Sub cmbSave_Click()
  Call SaveRec(Me)
End Sub

Private Sub cmbStatus_Click()
  If Me.NewRecord = True Then
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    Call ShowStatusForm(strCallingForm, True)
  Else
    Call ShowStatusForm(strCallingForm)
  End If
End Sub

Private Sub Form_AfterInsert()   ' After addition refresh list of combo box
  If Me.OpenArgs = cstDealForm Then Forms(cstDealForm)!cbxDealIssuer.Requery
End Sub

Private Sub Form_Current()
  Dim strFilter As String
  Dim lngCurrentDealNum As Long
  
  If Me.NewRecord = True Then Me.Modal = True
  If IsLoaded(cstFinStatusForm) Then 'show stats per current issuer
    If Forms!frmFinStat.Visible = True Then
      If IsNull(Me!lngIssuerNum) Then '
         strFilter = "[lngFinIssuerNum] = 1000"
      Else
         strFilter = "[lngFinIssuerNum] = " & Me!lngIssuerNum
      End If
      Forms!frmFinStat.FilterOn = True
      Forms!frmFinStat.Filter = strFilter
      Me.SetFocus
    End If
  End If
End Sub

Private Sub Form_Load()
  DoCmd.MoveSize 1440 * 5.5
  strCallingForm = CheckFormArgs(Me)
End Sub
