VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmNEWMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
'
' CLSMDL frmNEWMainForm Mar-03-16 11:25
'
'================================================================================
Private Sub cbxSelComp_AfterUpdate()
  Call MainFormSelComp_AfterUpdate
End Sub

Private Sub cbxSelComp_Change()
  lngPrevCo = Me!cbxSelComp.OldValue
  blnAddForecast = False
End Sub

Private Sub cbxSelComp_NotInList(NewData As String, Response As Integer)
  On Error GoTo cbxSelComp_NotInListErr
  MsgBox UCase(NewData) & " is not in the list." & vbCrLf & _
           "To enter a new Company, use the New button on the Companies page.", _
           vbExclamation, "Unknown Company"
  Response = acDataErrContinue
  Me!cbxSelComp.Undo

cbxSelComp_NotInListExit:
  Exit Sub
  
cbxSelComp_NotInListErr:
  MsgBox Err.Number & ": " & Err.Description, , cstFRM_Main & " :cbxSelComp_NotInList"
  Resume cbxSelComp_NotInListExit
End Sub

Private Sub cbxSelForecast_AfterUpdate()
  Call MainFormSelForecast_AfterUpdate
End Sub

Private Sub cbxSelForecast_Change()
  lngPrevFrc = Me!cbxSelForecast.OldValue
End Sub

Private Sub cbxSelForecast_NotInList(NewData As String, Response As Integer)
  On Error GoTo cbxSelForecast_NotInListErr
  MsgBox UCase(NewData) & " is not in the list." & vbCrLf & _
           "To enter a new Forecast, use the New button on the Forecast List page.", _
           vbExclamation, "Unknown Forecast"
  Response = acDataErrContinue
  Me!cbxSelForecast.Undo

cbxSelForecast_NotInListExit:
  Exit Sub
  
cbxSelForecast_NotInListErr:
  MsgBox Err.Number & ": " & Err.Description, , cstFRM_Main & " :cbxSelForecast_NotInList"
  Resume cbxSelForecast_NotInListExit
End Sub

Private Sub Form_Activate()
  DoCmd.Maximize
End Sub

Private Sub Form_Close()
  If Not blnExcelAlreadyRunning Then 'PFM started XL
    Excel.Application.Visible = True
    Excel.Application.Quit
  End If

  If Not blnWordAlreadyRunning Then
    Word.Application.Visible = True
    Word.Application.Activate
    Word.Application.Quit
  End If
  If IsLoaded(cstFRM_Closing) Then DoCmd.Close acForm, cstFRM_Closing
End Sub

Private Sub Form_Load()
  DoCmd.Hourglass True
  blnNewCo = False: blnExcelAlreadyRunning = False
  Me!txtCurBtn = "": Me!txtPrevBtn = ""
  Me!dteQEndDate = Date
  Me!lblDefBud.Visible = False
  Call ChangeAndToggleMainBtn(0)
  DoCmd.RunMacro "ResetSubForm0Folio"
  'Call ChangeAndToggleMainBtn(0)
  DoCmd.Hourglass False
End Sub

Private Sub lblBtn1_Click()
  If Not blnAddForecast Then Call ResetSubForm1Frcst
End Sub

Private Sub lblClosing_Click()
  Call OpenClosingInfo
End Sub

Private Sub lblExit_Click()
  DoCmd.RunCommand acCmdAppMinimize
  'Reset subform as unbound:
  Call ChangeAndToggleMainBtn(0)
  Me!sfrmAny.SourceObject = Null    'cstSFRM_Folio
  
  'If (IsLoaded(cstFRM_Fields)) Then DoCmd.Close acForm, cstFRM_Fields
  'If (IsLoaded(cstFRM_Closing)) Then DoCmd.Close acForm, cstFRM_Closing

  DoCmd.Quit acQuitSaveAll ' acExit
End Sub
