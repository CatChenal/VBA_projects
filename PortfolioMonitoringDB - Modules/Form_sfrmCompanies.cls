VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmCompanies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
'
' CLSMDL sfrmCompanies Mar-10-03 16:05
' Prev: Nov-20-02 14:45
'
'================================================================================

Private Sub dteRealized_DblClick(Cancel As Integer)
  Call OpenCalendar(Me, "dteRealized")
End Sub

Private Sub Form_Close()
  If Me.Dirty Then lblSave_Click
End Sub

Private Sub Form_Current()
  Dim frmFlds As Form
  
  If IsLoaded(cstFRM_Cal) Then DoCmd.Close acForm, cstFRM_Cal
  With Me
    blnNewCo = .NewRecord
    Call AvailableButtons(blnNewCo) '(Reset labels)
    .AllowAdditions = blnNewCo
    
    If Not blnNewCo Then
      !txtTotRecs = "Total: " & .RecordsetClone.RecordCount
      lngCurrentComp = !lngCompID
      If Not IsNull(.Parent!cbxSelComp) Then
        If .Parent!cbxSelComp <> lngCurrentComp Then
          .Parent!txtDefaultFRCDesc = ""
          .Parent!lblDefBud.Visible = False
          .Parent!cbxSelComp = lngCurrentComp    'update main form cbx
          .Parent!cbxSelForecast.Requery
          Call MainFormSelForecast_AfterUpdate
        End If
      Else
        .Parent!cbxSelComp = lngCurrentComp
      End If
    
      Call SyncFieldsFormRecs(False)
    End If
  End With
  DoCmd.Hourglass False
  Set frmFlds = Nothing
  
End Sub

Private Sub Form_Load()
  Me!blnRealized = False
  If lngCurrentComp = 0 Then lngCurrentComp = Me.Parent!cbxSelComp
  Call SyncSubRecord("lngCompID", lngCurrentComp)
  If blnNewCo Then
    Call lblNew_Click    'blnNewCo has been set by the parent form; add a record
  End If
  DoCmd.Hourglass False
End Sub

Private Sub intRealized_AfterUpdate()
  With Me
    If !intRealized <> 0 Then
      .Parent!cbxSelComp.Requery
    Else
      !dteRealized = Null
    End If
    .Requery
    !txtTotRecs = "Total: " & .RecordsetClone.RecordCount
  End With
End Sub

Private Sub lblCancel_Click()
  DoCmd.Hourglass True
  With Me
    If .Dirty Then
      .Undo
      If blnNewCo Then
        If .RecordsetClone.RecordCount > 0 Then DoCmd.GoToRecord , , acPrevious
        GoTo lblCancel_ClickExit
      End If
    End If
    .DataEntry = False
    If !blnRealized Then !blnRealized = Not !blnRealized
  End With
lblCancel_ClickExit:
  DoCmd.Hourglass False
End Sub

Private Sub lblCancel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblCancel")
End Sub

Private Sub lblCancel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblCancel")
End Sub

Private Sub lblFields_Click()
  Call SyncFieldsFormRecs(True)
End Sub

Private Sub lblFields_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblFields")
End Sub

Private Sub lblFields_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblFields")
End Sub

Private Sub lblMarkRealized_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblMarkRealized")
End Sub

Private Sub lblMarkRealized_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblMarkRealized")
End Sub

Private Sub lblNew_Click()
  Dim lngID As Long

  With Me
    If Not .Dirty Then
      .AllowAdditions = True
      .DataEntry = True
      blnNewCo = True
      Call AvailableButtons(True)
    End If
  End With
End Sub

Private Sub lblMarkRealized_Click()
  If Me!intRealized = 1 Then Exit Sub
  Call lblViewAll_Click
  Me!intRealized = 1
  Call AvailableButtons(True)
  'Call OpenCalendar(Me, "dteRealized")
End Sub

Private Sub lblNew_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblNew")
End Sub

Private Sub lblNew_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblNew")
End Sub

Private Sub lblSave_Click() 'sfrmCompanies
  Dim strCo As String
  
  On Error GoTo lblSave_ClickErr
  With Me
    If .Dirty Then
     If MsgBox("Save current record?", vbQuestion + vbYesNo, "Updated record") = vbYes Then
        strCo = .txtName
        DoCmd.RunCommand acCmdSaveRecord
        .DataEntry = False
      
        lngCurrentComp = !lngCompID
        .Parent!cbxSelComp.Requery  'refresh main form's combo box
        .Parent!cbxSelComp = lngCurrentComp
        
        Call CreateReviewFolder(.Parent!cbxSelComp.Column(1))
        Call SetNewCoFields(lngCurrentComp)   ': will add required fields in tblCompFields
      Else
        Call lblCancel_Click
      End If
    Else
      .Parent!cbxSelComp.Requery  'refresh main form's combo box
      .Parent!cbxSelComp = lngCurrentComp
      Call lblCancel_Click
    End If
  End With
  
lblSave_ClickExit:
  Exit Sub
  
lblSave_ClickErr:
  MsgBox "Error(" & Err.Number & ": " & Err.Description, vbExclamation, "Button: lblSave_Click"
  Resume lblSave_ClickExit
End Sub

Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblSave")
End Sub

Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblSave")
End Sub

Private Sub lblViewAll_Click()
  Dim str$
  DoCmd.Hourglass True
  With Me
    str$ = !lblViewAll.Caption ' Swap current caption with tag value & update hidden fld.
    !lblViewAll.Caption = !lblViewAll.Tag
    !lblViewAll.Tag = str$
    !blnRealized = Not !blnRealized   ' Toggle flag value to what it was not
    Call RefreshSet(Me!lngCompID)        ' Enable/disable filter accordingly
  End With
  DoCmd.Hourglass False
End Sub

Sub RefreshSet(lngCoID As Long)
  Dim str$
  
  With Me
    .FilterOn = Not !blnRealized 'filter is on when realized deals are not shown
    If Not !blnRealized Then
      str$ = "[tblCompanies]![intRealized] = 0"
      .Filter = str$
    End If
    Call AvailableRealizedFields
    !txtTotRecs = "Total: " & .RecordsetClone.RecordCount
    Call SyncSubRecord("lngCompID", lngCoID)  '.Parent!cbxSelComp)
  End With
  
End Sub

Private Sub lblViewAll_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblViewAll")
End Sub

Private Sub lblViewAll_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblViewAll")
End Sub

Private Sub txtName_AfterUpdate()
  Me!txtName = UCase(Me!txtName)
End Sub

Sub AvailableButtons(blnNewRecOrNewView As Boolean)
  With Me
    .Cycle = Abs(blnNewRecOrNewView)
    !lblFields.Visible = Not blnNewRecOrNewView
    !lblMarkRealized.Visible = Not blnNewRecOrNewView
    !lblViewAll.Visible = Not blnNewRecOrNewView
    !lblNew.Visible = Not blnNewRecOrNewView
    !txtTotRecs.Visible = Not blnNewRecOrNewView
    !lblCancel.Visible = blnNewRecOrNewView
    !lblSave.Visible = blnNewRecOrNewView
  End With
  Call AvailableRealizedFields
End Sub

Sub AvailableRealizedFields()
  With Me
    !lblRealized.Visible = !blnRealized
    !lblDateRealized.Visible = !blnRealized
    !intRealized.Visible = !blnRealized
    If Not !blnRealized Then !txtName.SetFocus  ' Change focus before hiding
    !dteRealized.Visible = !blnRealized
  End With
End Sub

Sub SyncFieldsFormRecs(blnShowFieldsForm As Boolean)
  Dim frm As Form
  Dim intView As Integer
  intView = 0 'normal; Form is modal if visible
  DoCmd.Hourglass True
  
  If IsLoaded(cstFRM_Fields) Then
    strFieldsFormFilter = "[tblCompFields]![lngCompID]=" & lngCurrentComp
    Set frm = Forms(cstFRM_Fields)
    With frm
      .Filter = strFieldsFormFilter
      .FilterOn = True
      If Not .Visible Then .Visible = blnShowFieldsForm 'keep visible if already visible
    End With
    Set frm = Nothing
  Else
    If Not blnShowFieldsForm Then intView = acHidden
    DoCmd.OpenForm cstFRM_Fields, , , strFieldsFormFilter, , intView
  End If
  With Forms(cstFRM_Fields)
    !cbxCompID = lngCurrentComp
    !cbxCompID.Requery
    !cbxFldID.Requery
    .Requery
    !txtTotRecs = "Total: " & .RecordsetClone.RecordCount
    .Modal = .Visible
  End With
  
  DoCmd.Hourglass False
End Sub
