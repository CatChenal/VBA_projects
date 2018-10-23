VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
'
' CLSMDL sfrmSummary Mar-10-03 12:30
' Prev  Nov-20-02 14:45
'
'================================================================================
Const cstDiscCheck = "PROBLEM: This Company has discontinuous series."

Private Sub cbxSelPeriodEnd_AfterUpdate()
  Dim blnWarn As Boolean
  'On Error GoTo cbxSelPeriodEnd_AfterUpdateErr
  
  With Me
    If !cbxSelPeriodEnd <> CDate(0) Then
      If !lblTitlePeriod.Caption = !cbxSelPeriodEnd.Column(1) Then
        If MsgBox("You have selected the same date as before: " & vbCrLf & _
                  "Do you really want to recalculate?", _
                  vbQuestion + vbYesNo, "Identical Date") = vbNo Then
          GoTo cbxSelPeriodEnd_AfterUpdateExit
        End If
      End If
      !lblTitlePeriod.Caption = !cbxSelPeriodEnd.Column(1)
      !lblWarn.Visible = True
      !lblWarn.Caption = ""
      !lblWarn.Visible = False
           
      If lngCurrentComp = 0 Then lngCurrentComp = .Parent!cbxSelComp
      If lngCurrentForecast = 0 Then lngCurrentForecast = .Parent!cbxSelForecast
      ' Check discontinuity of all company forecasts (first  arg=0):
      blnWarn = CheckDiscontinuity(0, lngCurrentComp, False)
      If blnWarn Then
        !lblWarn.Visible = True
        !lblWarn.Caption = cstDiscCheck
        GoTo cbxSelPeriodEnd_AfterUpdateExit
      End If
            
      Call RefreshSummaryGrid
            
    Else
      beep
      !lblWarn.Caption = "PROBLEM: Invalid period end date."
    End If
  End With
  
cbxSelPeriodEnd_AfterUpdateExit:
  Me!ocxFlxGridForm.MousePointer = flexDefault
  Exit Sub
  
cbxSelPeriodEnd_AfterUpdateErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , "cbxSelPeriodEnd_AfterUpdate"
  Resume cbxSelPeriodEnd_AfterUpdateExit
End Sub

Private Sub Form_Load()
  Dim str As String
  
  DoCmd.Hourglass True
  With Me
    !lblTitlePeriod.Caption = !lblTitlePeriod.Tag 'default value
    If blnNoFrc Then
       str = "<<no forecast>>"
       .Parent!lblDefBud.Visible = False
     Else
       str = .Parent!cbxSelForecast.Column(1)
       .Parent!lblDefBud.Visible = Nz(.Parent!cbxSelForecast.Column(3), 0)
     End If
    !txtCurrentCompany = .Parent!cbxSelComp.Column(1)
    !txtSmryTitle = str & " vs. " & .Parent!txtDefaultFRCDesc
  End With
  DoCmd.Hourglass False
End Sub

Private Sub cbxSelPeriodEnd_BeforeUpdate(Cancel As Integer)
  On Error GoTo cbxSelPeriodEnd_BeforeUpdateErr
  
  With Me
    If IsNull(!cbxSelPeriodEnd) Then
      beep
      !lblTitlePeriod.Caption = !lblTitlePeriod.Tag 'default value
      !cbxSelPeriodEnd.Undo
      !cbxSelPeriodEnd.SetFocus
      Cancel = True
    End If
  End With

cbxSelPeriodEnd_BeforeUpdateExit:
  Exit Sub
  
cbxSelPeriodEnd_BeforeUpdateErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , "cbxSelPeriodEnd_BeforeUpdate"
  Resume cbxSelPeriodEnd_BeforeUpdateExit
End Sub

Private Sub lblExport_Click()
  DoCmd.Hourglass True
  If Me!ocxFlxGridForm.Rows < 5 Or Me!cbxSelPeriodEnd = CDate(0) Then
    beep
    Exit Sub
  End If
  Call CopyGridDataToXLT(Me)
  DoCmd.Hourglass False
End Sub

Private Sub lblExport_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblExport")
End Sub

Private Sub lblExport_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblExport")
End Sub

Private Sub lblFields_Click()
  DoCmd.Hourglass True
  lngCurrentComp = Me.Parent!cbxSelComp
  DoCmd.OpenForm cstFRM_Fields, , , "[lngCompID]=" & lngCurrentComp
  Forms(cstFRM_Fields)!cbxCompID = lngCurrentComp
  DoCmd.Hourglass False
End Sub

Private Sub lblFields_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblFields")
End Sub

Private Sub lblFields_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblFields")
End Sub
