VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'===============================================================================
' DBFrontEnd
' CLSMDL frmManagement May-10-02 16:10
'
'===============================================================================
Const cstThisForm = cstMgmtForm

Private Sub cbxSelCompany_BeforeUpdate(Cancel As Integer)
  If Me.Dirty Then
    Me!cbxSelCompany.Undo
    Me!cbxSelCompany = Me!cbxSelCompany.OldValue
    Cancel = True
    MsgBox "The current record has not been saved: cannot proceed", vbExclamation, "Company Selection Change"
  End If
End Sub

Private Sub cbxSelCompany_AfterUpdate()
  Dim lngCo As Long
  
  lngCo = Me!cbxSelCompany.Column(0)
  Call CheckMgmt(lngCo)
  
End Sub

Private Sub cmbGetOLAdrs_Click()
  On Error GoTo cmbGetOLAdrsErr
  DoCmd.Hourglass True
  If (Len(Me!txtMgrLast & "") > 0 And Len(Me!txtMgrFirst & "") > 0) Then
    Call GetOLData(Me, Me!txtMgrLast, Me!txtMgrFirst, "Address")
  Else
    MsgBox "Both First and Last names are required.", vbExclamation, cstThisForm & "cmbGetOLAdrs"
  End If
  DoCmd.Hourglass False
  
cmbGetOLAdrsExit:
  Exit Sub

cmbGetOLAdrsErr:
  Call EndOutlook
  DoCmd.Hourglass False
  MsgBox "Error: (" & Err & ") " & Err.Description, vbExclamation, cstThisForm & "cmbGetOLAdrs"
  Resume cmbGetOLAdrsExit
End Sub

Private Sub cmbGetOLEmail_Click()
  On Error GoTo cmbGetOLEmailErr
  DoCmd.Hourglass True
  If (Len(Me!txtMgrLast & "") > 0 And Len(Me!txtMgrFirst & "") > 0) Then
    Call GetOLData(Me, Me!txtMgrLast, Me!txtMgrFirst, "Email/Phones")
  Else
    MsgBox "Both First and Last names are required.", vbExclamation, cstThisForm & "cmbGetOLEmail"
  End If
  DoCmd.Hourglass False
  
cmbGetOLEmailExit:
  Exit Sub

cmbGetOLEmailErr:
  Call EndOutlook
  DoCmd.Hourglass False
  MsgBox "Error: (" & Err & ") " & Err.Description, vbExclamation, cstThisForm & "cmbGetOLEmail"
  Resume cmbGetOLEmailExit
End Sub

Private Sub cmbNewRec_Click()
  Call NewRec(Me)
End Sub

Private Sub cmbSameCoNew_Click()
  Dim rstMgmtForm As Recordset
  Dim lngCoID As Long
  Dim varBkmk
  Const cst = "New Manager, Same Company"
  
  On Error GoTo cmbSameCoNewErr
  
  lngCoID = Me!lngCompanyID
  If Me.Dirty Then Call SaveRec(Me)
  
  DoCmd.Hourglass True
  varBkmk = Me.Bookmark
 
  Set rstMgmtForm = Me.RecordsetClone
  ' Copy the bookmarked record co & addrs fields to new rec: (let user update)
  With rstMgmtForm
    .AddNew
    !lngCompanyID = lngCoID
    !txtMgrFirst = "First"
    !txtMgrLast = "Last"
    !txtAddress1 = Me!txtAddress1
    !txtAddress2 = Me!txtAddress2
    !txtCity = Me!txtCity
    !txtState = Me!txtState
    !txtZip = Me!txtZip
    !txtCountry = Me!txtCountry
    
    If MsgBox("Copy Phone & Fax?", vbQuestion + vbYesNo, cst) = vbYes Then
      !txtMgrPhone = Me!txtMgrPhone
      !txtMgrFax = Me!txtMgrFax
    End If
    If MsgBox("Copy Email?", vbQuestion + vbYesNo, cst) = vbYes Then
      !txtMgrEmail = Me!txtMgrEmail
    End If
    .Bookmark = varBkmk
  End With
  DoCmd.Hourglass False
  Me!txtMgrFirst.SetFocus
  
cmbSameCoNewExit:
  Exit Sub
  
cmbSameCoNewErr:
  DoCmd.Hourglass False
  MsgBox "Error: (" & Err & ") " & Err.Description, vbExclamation, cstThisForm & "cmbSameCoNew"
  Resume cmbSameCoNewExit
End Sub

Private Sub cmbSaveRec_Click()
  DoCmd.OpenForm "frmUpdContDLG"
  strCallingForm = cstThisForm
End Sub

Private Sub cmbCloseForm_Click()
  Call CloseFrm(Me)
End Sub

Private Sub cmbUndoRec_Click()
  Call UndoRec(Me)
End Sub

Private Sub dteDateCASigned_DblClick(Cancel As Integer)
  Me!dteDateCASigned = Date
End Sub

Private Sub dteDateMet_DblClick(Cancel As Integer)
  Me!dteDateMet = Date
End Sub

Private Sub Form_Activate()
 DoCmd.Maximize
 strCallingForm = cstThisForm
End Sub

Private Sub Form_AfterUpdate()
  Me!cmbSameCoNew.Enabled = True
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
  Me!cmbSameCoNew.Enabled = False
End Sub

Private Sub Form_Current()
  Dim blnNewMgr As Boolean
  
  With Me
    blnNewMgr = CBool(.NewRecord)
    !lblNewMgr.Visible = blnNewMgr
    !cmbSameCoNew.Enabled = (Not blnNewMgr)
    'Reset header ctls: co number selection box & related controls:
    !lblSelCo.Visible = (Not blnNewMgr)
    !cbxSelCompany.Visible = (Not blnNewMgr)
    If Not blnNewMgr Then !cbxSelCompany = !lngCompanyID
  End With
End Sub

Private Sub Form_Open(Cancel As Integer)
  strCallingForm = CheckFormArgs(Me)
End Sub

Private Sub cmbGetOLData_Click()  'to overwrite all fields on form
  Dim lngErr  As Long
  On Error GoTo cmbGetOLDataErr
  
  DoCmd.Hourglass True
  If (Len(Me!txtMgrLast & "") > 0 And Len(Me!txtMgrFirst & "") > 0) Then
    Call GetOLData(Me, Me!txtMgrLast, Me!txtMgrFirst)
  Else
    MsgBox "Both First and Last names are required.", vbExclamation, cstThisForm & "cmbGetOLData"
  End If
  DoCmd.Hourglass False
  
cmbGetOLDataExit:
  Call EndOutlook
  Exit Sub

cmbGetOLDataErr:
  DoCmd.Hourglass False
  lngErr = CLng(Err.Number)
  MsgBox "Error: (" & lngErr & ") " & Err.Description, vbExclamation, cstThisForm & "cmbGetOLData"
  Resume cmbGetOLDataExit
End Sub

Private Sub txtCity_AfterUpdate()
  Me!txtCity.Requery
End Sub

Private Sub txtCountry_AfterUpdate()
  Me!txtCountry.Requery
End Sub

Private Sub txtState_AfterUpdate()
  Me!txtState.Requery
End Sub

Private Sub txtZip_AfterUpdate()
  Me!txtZip.Requery
End Sub

Public Sub CheckMgmt(lngCompID As Long)
 ' called after cbxSelCompany afterupd
  Dim rst As DAO.Recordset
  Dim strCrit As String, strMsg As String
  Dim intResponse As Integer
  
  On Error GoTo CheckMgmtErr
  
  DoCmd.Hourglass True

  strCrit = "[lngCompanyID] = " & lngCompID
  Set rst = Me.RecordsetClone
  With rst
    .FindFirst strCrit
    If Not .NoMatch Then
      Me.Bookmark = .Bookmark
      Me.SetFocus
    Else
      strMsg = "There is no Management information for the selected company:" & vbCrLf & _
             "Do you want to create a record?"
      intResponse = MsgBox(strMsg, vbQuestion + vbYesNo, "No Manager found")
      If intResponse = vbNo Then
        Me!cbxSelCompany.Undo
      Else
        DoCmd.GoToRecord acActiveDataObject, , acNewRec
        Me!lngCompanyID = lngCompID
      End If
    End If
  End With
  DoCmd.Hourglass False
  
CheckMgmtExit:
  Set rst = Nothing
  Exit Sub
    
CheckMgmtErr:
  DoCmd.Hourglass False
  MsgBox Err.Number & ": " & Err.Description, , cstThisForm & ": CheckMgmt"
  Resume CheckMgmtExit
End Sub
