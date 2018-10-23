VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFilterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'===============================================================================
' DBFrontEnd
' CLSMDL FilterForm Sep-05-03 14:40
'     Oct-10-02
'                   - Changed MouseDown macro to proc so that the actual
'                   button is checked (macro was alwas using the right one).
'                   Also changed other MouseDown procs with same effect.
'                   Aug-27-02 11:15
'                   - Added txtEditOrder to enable Undo from the most
'                   recently edited ctl.
'                   - Modified Cancel btn so that latest changes are undone.
'                   Prev: Jun-14-02 16:00
'
' Because the form is left opened(hidden) after use, it is only oened once
' so the .OpenArgs method can not be reused. The hidden field tbxCallingForm
' is used as a workaround.
'
'===============================================================================
Const cstMDL = cstFilterForm

Private Sub cbxSelAnalyst_AfterUpdate()
  Call AddToEditList(Screen.ActiveControl.TabIndex)
End Sub

Private Sub cbxSelAnalyst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
End Sub

Private Sub cbxSelControl_AfterUpdate()
  Call AddToEditList(Screen.ActiveControl.TabIndex)
End Sub

Private Sub cbxSelControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
End Sub

Private Sub cbxSelDisposition_AfterUpdate()
  Call AddToEditList(Screen.ActiveControl.TabIndex)
End Sub

Private Sub cbxSelDisposition_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
End Sub

Private Sub cbxSelIndustry_AfterUpdate()
  Call AddToEditList(Screen.ActiveControl.TabIndex)
End Sub

Private Sub cbxSelIndustry_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
End Sub

Private Sub cbxSelIssuer_AfterUpdate()
  Call AddToEditList(Screen.ActiveControl.TabIndex)
End Sub

Private Sub cbxSelIssuer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
End Sub

Private Sub cbxSelSource_AfterUpdate()
  Call AddToEditList(Screen.ActiveControl.TabIndex)
End Sub

Private Sub cbxSelSource_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
End Sub

Private Sub cbxSelSourceType_AfterUpdate()
  Call AddToEditList(Screen.ActiveControl.TabIndex)
End Sub

Private Sub cbxSelSourceType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
End Sub

Private Sub cbxSelSponsor_AfterUpdate()
  Call AddToEditList(Screen.ActiveControl.TabIndex)
End Sub

Private Sub cbxSelSponsor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
End Sub

Private Sub Form_GotFocus()
  Call ApplyFormFilter(cstMainForm)
End Sub

Private Sub Form_Load()
  Me!tbxCallingForm = cstMainForm
  Call ResetFilterWindow(cstDefWidth, cstDefHeight)
End Sub

Private Sub cbxSelYear1_AfterUpdate() 'uses ctl tab index
  Call DateCtlAfterUpdate(1)
End Sub

Private Sub cbxSelYear2_AfterUpdate()
  Call DateCtlAfterUpdate(1)
End Sub

Private Sub cbxSelMonth1_AfterUpdate()
  Call DateCtlAfterUpdate(2)
End Sub

Private Sub cbxSelMonth2_AfterUpdate()
  Call DateCtlAfterUpdate(2)
End Sub

Private Sub cbxSelQtr1_AfterUpdate()
  Call DateCtlAfterUpdate(3)
End Sub

Private Sub cbxSelQtr2_AfterUpdate()
  Call DateCtlAfterUpdate(3)
End Sub

Private Sub Form_LostFocus()
  Call ApplyFormFilter(cstMainForm)
End Sub

Private Sub tbxSelDate1_AfterUpdate()
  Call DateCtlAfterUpdate(4)
End Sub

Private Sub tbxSelDate2_AfterUpdate()
  Call DateCtlAfterUpdate(4)
End Sub

Private Sub cbxSelYear1_BeforeUpdate(Cancel As Integer)
  Cancel = DateRangeBeforeUpdate("start", "year", "cbxSelYear1")
End Sub

Private Sub cbxSelYear2_BeforeUpdate(Cancel As Integer)
  Cancel = DateRangeBeforeUpdate("end", "year", "cbxSelYear2")
End Sub

Private Sub cbxSelMonth1_BeforeUpdate(Cancel As Integer)
  Cancel = DateRangeBeforeUpdate("start", "month", "cbxSelMonth1")
End Sub

Private Sub cbxSelMonth2_BeforeUpdate(Cancel As Integer)
  Cancel = DateRangeBeforeUpdate("end", "month", "cbxSelMonth2")
End Sub

Private Sub cbxSelQtr1_BeforeUpdate(Cancel As Integer)
  Cancel = DateRangeBeforeUpdate("start", "quarter", "cbxSelQtr1")
End Sub

Private Sub cbxSelQtr2_BeforeUpdate(Cancel As Integer)
  Cancel = DateRangeBeforeUpdate("end", "quarter", "cbxSelQtr2")
End Sub

Private Sub tbxSelDate1_BeforeUpdate(Cancel As Integer)
  Cancel = DateRangeBeforeUpdate("start", "date", "tbxSelDate1")
End Sub

Private Sub tbxSelDate2_BeforeUpdate(Cancel As Integer)
  Cancel = DateRangeBeforeUpdate("end", "date", "tbxSelDate2")
End Sub

Private Sub cbxSelYear1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
  Call cbxSelYear1_AfterUpdate
End Sub

Private Sub cbxSelYear2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
  Call cbxSelYear2_AfterUpdate
End Sub

Private Sub cbxSelMonth1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
  Call cbxSelMonth1_AfterUpdate
End Sub

Private Sub cbxSelMonth2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
  Call cbxSelMonth2_AfterUpdate
End Sub

Private Sub cbxSelQtr1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
  Call cbxSelQtr1_AfterUpdate
End Sub

Private Sub cbxSelQtr2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
  Call cbxSelQtr2_AfterUpdate
End Sub

Private Sub tbxSelDate1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
  Call tbxSelDate1_AfterUpdate
End Sub

Private Sub tbxSelDate2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ClearCtlWithRightMouseDown(Screen.ActiveControl, Button)
  Call tbxSelDate2_AfterUpdate
End Sub

Private Sub tbxSelDate1_DblClick(Cancel As Integer)
  Call OpenCalendar(Me, "tbxSelDate1")
End Sub

Private Sub tbxSelDate2_DblClick(Cancel As Integer)
  Call OpenCalendar(Me, "tbxSelDate2")
End Sub

'=== btns events 1 ==========================
Private Sub cmbCancel_Click()
  Call UndoLastEdits
End Sub

Private Sub cmbReset_Click()
  Call ResetFilterForm
End Sub

Private Sub cmbOK_Click() ' caption= APPLY
  With Me
    If !chkYear Then
      If IsNull(!cbxSelYear1) And IsNull(!cbxSelYear2) Then !chkYear = 0
    End If
    
    If !chkMonth Then
      If IsNull(!cbxSelMonth1) And IsNull(!cbxSelMonth2) Then
        !chkMonth = 0
        Call DateSelectionCheckboxClick(2)
      End If
    End If
    
    If !chkQuarter Then
      If IsNull(!cbxSelQtr1) And IsNull(!cbxSelQtr2) Then
        !chkQuarter = 0
        Call DateSelectionCheckboxClick(3)
      End If
    End If
    
    If !chkRange Then
      If IsNull(!tbxSelDate1) And IsNull(!tbxSelDate2) Then
        !chkRange = 0
        Call DateSelectionCheckboxClick(4)
      End If
    End If
  End With
  
  Call GetFilteredRecords(Me!tbxCallingForm) 'clears the edit order ctl
  
  Me.Visible = False
End Sub
'=== btns events 0 ==========================

Private Sub chkYear_Click()
  Call DateSelectionCheckboxClick(1)
End Sub

Private Sub chkMonth_Click()
  Call DateSelectionCheckboxClick(2)
End Sub

Private Sub chkQuarter_Click()
  Call DateSelectionCheckboxClick(3)
End Sub

Private Sub chkRange_Click()
  Call DateSelectionCheckboxClick(4)
End Sub


