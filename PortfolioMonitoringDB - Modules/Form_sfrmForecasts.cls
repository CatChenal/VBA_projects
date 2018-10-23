VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmForecasts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
'
' CLSMDL sfrmForecasts Jul-31-03 11:20
'
'================================================================================
Dim blnNewForecast As Boolean
Dim lngPrevFRCType As Long

Private Sub cbxForecastTypeID_AfterUpdate()
  With Me
    If IsNull(!cbxForecastTypeID) Then Exit Sub
    
    If Not IsDate(!dteForecastDate) Then
      !txtForecastDesc = !cbxCompID.Column(1) & " - " & !cbxForecastTypeID.Column(1)
    Else
      !txtForecastDesc = !cbxCompID.Column(1) & " - " & !cbxForecastTypeID.Column(1) & _
                                                " - " & !dteForecastDate
    End If
  End With
End Sub

Private Sub cbxForecastTypeID_Enter()
  lngPrevFRCType = Nz(Me!lngForecastTypeID, 0)
End Sub

Private Sub dteForecastDate_AfterUpdate()
  With Me
    If IsNull(!cbxForecastTypeID) Then
      !txtForecastDesc = !cbxCompID.Column(1) & " - " & !dteForecastDate
    Else
      !txtForecastDesc = !cbxCompID.Column(1) & " - " & !cbxForecastTypeID.Column(1) & _
                                                " - " & !dteForecastDate
    End If
    SendKeys ("{TAB}")
  End With
End Sub

Private Sub dteForecastDate_DblClick(Cancel As Integer)
  Call OpenCalendar(Me, "dteForecastDate")
End Sub

Private Sub Form_Current()
  blnNewForecast = False
  If IsLoaded(cstFRM_Cal) Then DoCmd.Close acForm, cstFRM_Cal
  
  With Me
    blnNewForecast = .NewRecord
    .AllowAdditions = blnNewForecast
    .AllowEdits = blnNewForecast
    ' Reset labels:
    Call AvailableButtons("", blnNewForecast)
    
    If Not blnNewForecast Then
      lngCurrentComp = !cbxCompID
      lngCurrentForecast = Nz(!lngForecastID, 0)
      
      .Parent!cbxSelComp = lngCurrentComp
      .Parent!cbxSelForecast.Requery
      .Parent!cbxSelForecast = lngCurrentForecast
      .Parent!lblDefBud.Visible = Nz(.Parent!cbxSelForecast.Column(3), 0)  'intDefaultBud
      .Parent!txtDefaultFRCDesc = GetCoDefaultBudgetDesc(lngCurrentComp)
      If lngPrevCo <> lngCurrentComp Then Call Validate
    End If
  End With
  lngPrevCo = lngCurrentComp 'reset with curr comp
End Sub

Private Sub Form_Load()
  Call SyncSubRecord("lngCompID", lngCurrentComp, "lngForecastID", lngCurrentForecast)
  lngPrevCo = Me!cbxCompID
  DoCmd.Hourglass False
End Sub

Private Sub intDefaultBudget_BeforeUpdate(Cancel As Integer)
  Dim dbDAO As DAO.Database
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim iCount As Integer
  iCount = 0
  On Error GoTo intDefaultBudget_BeforeUpdateErr

  If Me!lngForecastTypeID <> 2 Then Exit Sub

  Set dbDAO = CurrentDb
  Set qdf = dbDAO.QueryDefs("qryDefBudgetCount")
 If lngCurrentComp = 0 Then lngCurrentComp = Me.Parent!cbxSelComp

  qdf.Parameters(0) = lngCurrentComp
  Set rst = qdf.OpenRecordset
  If rst.AbsolutePosition <> -1 Then
    iCount = rst(0)
  End If
  rst.Close
  Set rst = Nothing
  qdf.Close
  Set qdf = Nothing

  With Me
    If iCount <> 0 And !intDefaultBudget Then 'dup
        MsgBox "You cannot flag this forecast to a Default Budget:" & vbCrLf & _
               "Another Budget forecast has already been flagged." & vbCrLf & vbCrLf & _
               "To switch default budget, first deselect the one" & vbCrLf & _
               "already defined.", vbExclamation, "Default Budget Forecasts Are Unique"
        !intDefaultBudget.Undo
        Cancel = True
    Else

      If iCount = 1 And Not !intDefaultBudget Then
        If MsgBox("You are about to deselect this Default Budget:" & vbCrLf & _
                  "If you do not select another one, no budget data will" & vbCrLf & _
                  "appear on the Summary Form." & vbCrLf & _
                  "Is this OK?.", vbQuestion + vbYesNo, _
                  "Change Default Budget?") = vbNo Then
          !intDefaultBudget.Undo
          Cancel = True
        End If
      End If

    End If
  End With

intDefaultBudget_BeforeUpdateExit:
  Set rst = Nothing
  Set qdf = Nothing
  Set dbDAO = Nothing
  Exit Sub

intDefaultBudget_BeforeUpdateErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, "intDefaultBudget_BeforeUpdate"
  Resume intDefaultBudget_BeforeUpdateExit
End Sub

Private Sub lblCancel_Click()
  DoCmd.Hourglass True
  
  With Me
    If .Dirty Then
      .Undo
      If blnNewForecast Then
        If .RecordsetClone.RecordCount > 0 Then DoCmd.GoToRecord , , acPrevious
        .DataEntry = False
        .AllowAdditions = False
        blnNewForecast = False
        GoTo lblCancel_ClickExit
      End If
    End If
    .DataEntry = False
    .AllowEdits = False
    blnAddForecast = False
    Call AvailableButtons("", False)
  End With

lblCancel_ClickExit:
  DoCmd.Hourglass False
End Sub

Private Sub lblDefBudget_Click()
  MsgBox "A check mark indicates that the selected budget is used on the Summary Form.", _
         vbInformation, "Explanation: Default Budget"
End Sub

Private Sub lblDelete_Click()
  If MsgBox("Do you want to delete this forecast?" & vbCrLf & vbCrLf & _
            "Forecast: " & Me.Parent!cbxSelForecast.Column(1) & vbCrLf, _
            vbQuestion + vbYesNo, "Forecast Deletion") = vbYes Then
    Me.AllowDeletions = True
    DoCmd.RunCommand acCmdDeleteRecord
    Me.Requery
    Me.AllowDeletions = False
  End If
End Sub

Private Sub lblDelete_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblDelete")
End Sub

Private Sub lblDelete_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblDelete")
End Sub

Private Sub lblEdit_Click()
  With Me
     Call AvailableButtons("lblEdit", False)
    .AllowEdits = True
    !cbxCompID.SetFocus
    !cbxCompID = !cbxCompID
  End With
End Sub

Private Sub lblEdit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblEdit")
End Sub

Private Sub lblEdit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblEdit")
End Sub

Private Sub lblExport_Click()
  Call OpenExportForm
End Sub

Private Sub lblExport_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblExport")
End Sub

Private Sub lblExport_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblExport")
End Sub

Private Sub lblNew_Click()
  Call AddForecast
End Sub

Private Sub lblSave_Click() 'sfrmForecasts
  Dim bln As Boolean
  
  On Error GoTo lblSave_ClickErr
  DoCmd.Hourglass True
  
  With Me
    If .Dirty Then
      bln = blnNewForecast 'save locally, blnNewForecast is reset by me.newrecord prop after save
      If MsgBox("Save current record?", vbQuestion + vbYesNo, "Updated record") = vbYes Then
        DoCmd.RunCommand acCmdSaveRecord
        If bln Then Call CheckForecastType
        .DataEntry = False
        .AllowAdditions = False
        .AllowEdits = False
        blnAddForecast = False
      Else
        Call lblCancel_Click
      End If
    End If
  End With
  
lblSave_ClickExit:
  DoCmd.Hourglass False
  Exit Sub

lblSave_ClickErr:
  If Err = 3163 Then  'fld too small
    Me!txtForecastDesc = Me!cbxCompID.Column(1) & " - " & _
                         Me!cbxForecastTypeID.Column(1) & " (newest)"
    Resume Next
  Else
    If Err = 2113 Then
      MsgBox "Wrong Value Entered: " & Err.Description, , "Cannot Save"
    Else
      MsgBox "Error (" & Err.Number & "): " & Err.Description, , "lblSave_Click"
    End If
    Resume lblSave_ClickExit
  End If
End Sub

Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblSave")
End Sub

Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblSave")
End Sub

Private Sub txtForecastDesc_DblClick(Cancel As Integer)
  If Me.AllowEdits Then DoCmd.RunCommand acCmdZoomBox
End Sub

Sub CheckForecastType()
  Dim dbDAO As DAO.Database
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim strFRC As String, strNewDesc As String
  Dim bln As Boolean
  Dim r As Integer
  
  bln = False
  DoCmd.Hourglass True
  
  Set dbDAO = CurrentDb
  Set qdf = dbDAO.QueryDefs("qryCoForecastsOfSpecificType")
  ' The qry sorts the frc desc descendingly so the largest str is on top:
  ' this minimizes the number of passes.
  
  qdf.Parameters(0) = Me!cbxCompID
  qdf.Parameters(1) = Me!cbxForecastTypeID
  strNewDesc = Me!txtForecastDesc
  qdf.Parameters(2) = strNewDesc
  
  Set rst = qdf.OpenRecordset
  If rst.AbsolutePosition <> -1 Then
    rst.MoveLast
    rst.MoveFirst
    For r = 1 To rst.RecordCount - 1
      strFRC = rst!txtForecastDesc
      If Len(strFRC) >= Len(strNewDesc) Then
        strNewDesc = strFRC & " (last entered)"

        MsgBox "NOTE:" & vbCrLf & _
               "The forecast you have just entered had the same" & vbCrLf & _
               "description as another one of the same kind." & vbCrLf & _
               "It has been modified for better identification.", _
                vbInformation, "Description Amended"
        Me.AllowEdits = True
        Me!txtForecastDesc = strNewDesc
        Exit For
      End If
      rst.MoveNext
    Next r
  End If
  rst.Close
  qdf.Close
  Set rst = Nothing
  Set qdf = Nothing
  Set dbDAO = Nothing
  DoCmd.Hourglass False
End Sub

Sub AvailableButtons(strBtnClicked As String, blnNewFrc As Boolean)
  With Me
    Select Case strBtnClicked
      Case "lblEdit", "lblNew"
        !lblDelete.Visible = False
        !lblExport.Visible = False
        !lblNew.Visible = False
        !lblEdit.Visible = False
        !lblCancel.Visible = True
        !lblSave.Visible = True
      Case Else
        !lblDelete.Visible = Not blnNewFrc
        !lblExport.Visible = Not blnNewFrc
        !lblNew.Visible = Not blnNewFrc
        !lblEdit.Visible = Not blnNewFrc
        !lblCancel.Visible = blnNewFrc
        !lblSave.Visible = blnNewFrc
    End Select
  End With
End Sub
