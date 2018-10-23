VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDatasheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
' DBFrontEnd
' CLSMDL frmDatasheet Sep-05-03 14:40
' Prev: Sept-03-02 15:45
'
'================================================================================
Const cstThisForm = "frmDatasheet"
'-----------------------------------------------------------------------------

Private Sub cmbSelectRpt_Click()
  Forms(cstReportForm).Visible = True
End Sub

Private Sub cmbToList_Click()
  strCallingForm = cstThisForm
  Me.Visible = False
  If IsLoaded(cstFilterForm) Then
    Forms(cstFilterForm)!tbxCallingForm = strCallingForm
  End If
  Forms(cstMainForm).Visible = True
  Call RefreshActivateMainForm
End Sub

Private Sub Form_GotFocus()
  Call ApplyFormFilter(cstThisForm)
End Sub

Private Sub Form_Load()
  Call ResetToolbars
End Sub

Private Sub Form_Open(Cancel As Integer)
  If Not IsLoaded(cstMainForm) Then
    MsgBox "To view the result of a query, first make a selection via the " & vbCrLf & _
           "Report Selection form which is available from the Deal Selection Form.", _
           vbExclamation, "Form not directly accessible."
    Cancel = True
  End If
End Sub

Private Sub lblFilter_Click()
  strCallingForm = cstThisForm
  Call ShowFilterForm(cstThisForm, False)
End Sub
