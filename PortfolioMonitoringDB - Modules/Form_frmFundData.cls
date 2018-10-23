VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFundData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
'
' CLSMDL frmFFundData Apr-4-03 12:45
' Prev: Sept-04-02 16:00
'
'================================================================================
Private Sub cbxClosingID_AfterUpdate()
  Call RefreshFilter
  Call AllocCheck
End Sub

Private Sub Form_Close()
  If Me.Dirty Then DoCmd.RunCommand acCmdSave
End Sub

Private Sub Form_Open(Cancel As Integer)
  If Not IsLoaded(cstFRM_Closing) Then Cancel = True
  Call ResetClosingOrFundForm(cstFRM_Fund, CBool(Nz(Me.OpenArgs, 0)) And True)
  Call AllocCheck
End Sub

Private Sub lblNew_Click()
  Call ResetClosingOrFundForm(cstFRM_Fund, True)
End Sub

Private Sub lblSave_Click()
  On Error GoTo lblSave_ClickErr
  
  If Me.Dirty Then
    If MsgBox("Save record?", vbQuestion + vbYesNo, "Save Record") = vbYes Then
      DoCmd.Hourglass True
      DoCmd.RunCommand acCmdSaveRecord
      Call cbxClosingID_AfterUpdate
      Call ResetClosingOrFundForm(cstFRM_Fund, False)
    End If
  Else
    If Me!lblSave.Caption = "Close" Then DoCmd.Close
  End If
  
lblSave_ClickExit:
  DoCmd.Hourglass False
  Exit Sub
lblSave_ClickErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , "lblSave_Click"
  Resume lblSave_ClickExit
End Sub

Sub AllocCheck()
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim strMsg As String
  
  Set qdf = CurrentDb.QueryDefs("qryFundAllocSum")
  qdf.Parameters(0) = Me!cbxClosingID
  Set rst = qdf.OpenRecordset
  If rst.AbsolutePosition = -1 Then Exit Sub
  If Nz(rst(0), 1) <> 1 Then
    Me!cbxClosingID.SetFocus
    strMsg = "The total allocation for " & Me!cbxClosingID.Text & " does not equal 100%!" & vbCrLf & _
             "Please review and correct."
    MsgBox strMsg, vbExclamation, "Allocation Check"
  End If
End Sub

Sub RefreshFilter()
  Dim strFilter  As String
  On Error GoTo RefreshFilterErr
  
  With Me
    strFilter = "[tblFundData]![lngClosingID]=" & Forms(cstFRM_Closing)!lngClosingID
    .Filter = strFilter
    .FilterOn = True
    .Requery
  End With
RefreshFilterExit:
  Exit Sub
  
RefreshFilterErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, "frmFields:RefreshFilter"
  Resume RefreshFilterExit
End Sub
