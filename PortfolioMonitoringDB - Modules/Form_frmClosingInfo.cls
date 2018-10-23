VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmClosingInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
'
' CLSMDL cfrmClosingInfo Apr-04-03  15:45
'
'================================================================================

Private Sub cbxCompID_AfterUpdate()
  With Me
    lngCurrentComp = !cbxCompID
    If Not blnNewClosing Then .Requery
    ' If change of co, then synchronize calling form
    If Forms(cstFRM_Main)!cbxSelComp <> lngCurrentComp Then
      With Forms(cstFRM_Main)
        !cbxSelComp = lngCurrentComp
        !cbxSelComp.Requery
      End With
      If Not blnNewClosing Then Call MainFormSelComp_AfterUpdate
    End If
  End With
End Sub

Private Sub cbxCompID_BeforeUpdate(Cancel As Integer)
  If IsNull(DLookup("[lngCompID]", "tblClosingInfo", "[lngCompID]=" & Me!cbxCompID)) Then
    If MsgBox("There is no Closing information for " & Me!cbxCompID.Column(1) & ". " & vbCrLf & _
              "Do you want to enter it?", vbExclamation + vbYesNo, "No Closing Info") = vbYes Then
      blnNewClosing = True
    Else
      blnNewClosing = False
    End If
  End If
End Sub

Private Sub dteClosingDate_DblClick(Cancel As Integer)
  Me!dteClosingDate = Date
End Sub

Private Sub Form_Open(Cancel As Integer)
  If Not IsLoaded(cstFRM_Main) Then Cancel = True
  If IsNull(Forms(cstFRM_Main)!cbxSelComp) Then Cancel = True
  
  Me!cbxCompID = Forms(cstFRM_Main)!cbxSelComp
  Call ResetClosingOrFundForm(cstFRM_Closing, CBool(Nz(Me.OpenArgs, 0)) And True)
End Sub

Private Sub lblAlloc_Click()
  Dim blnNoAlloc As Boolean
  
  If (Me.Dirty Or IsNull(Me!lngClosingID)) Then
    MsgBox "You must save the Closing Information record before " & _
            "you can enter the related Fund Information", vbExclamation, "Save Current Record"
    Exit Sub
  End If
  blnNoAlloc = IsNull(DLookup("[lngClosingID]", "tblFundData", "[lngClosingID]=" & _
                              Forms(cstFRM_Closing)!lngClosingID))
  If blnNoAlloc Then 'ask
    If MsgBox("No Fund Allocation was entered for this Closing: do you want to enter it?", _
               vbQuestion + vbYesNo, "No Allocation") = vbNo Then
      Exit Sub
    End If
  End If
  DoCmd.OpenForm "frmFundData", , , , , , blnNoAlloc
End Sub

Private Sub lblAlloc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblAlloc")
End Sub

Private Sub lblAlloc_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblAlloc")
End Sub

Private Sub lblNew_Click()
  On Error GoTo blNew_ClickErr
  
  Call ResetClosingOrFundForm(cstFRM_Closing, True)
  Me!cbxCompID = lngCurrentComp
  
lblNewExit:
  Exit Sub
blNew_ClickErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , cstFRM_Closing & ": lblNew_Click"
  Resume lblNewExit
End Sub

Private Sub lblNew_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblNew")
End Sub

Private Sub lblNew_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblNew")
End Sub

Private Sub lblSave_Click()
  On Error GoTo lblSave_ClickErr
  
  If Me.Dirty Then
    If MsgBox("Save record?", vbQuestion + vbYesNo, "Save Record") = vbYes Then
      DoCmd.Hourglass True
      Me!lngCompID = Me!cbxCompID
      DoCmd.RunCommand acCmdSaveRecord
      blnNewClosing = False
      Me.Requery
      Call ResetClosingOrFundForm(cstFRM_Closing, False)
    End If
  Else
    If Me!lblSave.Caption = "Close" Then DoCmd.Close
  End If
  
lblSave_ClickExit:
  DoCmd.Hourglass False
  Exit Sub
lblSave_ClickErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , cstFRM_Closing & ": lblSave_Click"
  Resume lblSave_ClickExit
End Sub

Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblSave")
End Sub

Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblSave")
End Sub
