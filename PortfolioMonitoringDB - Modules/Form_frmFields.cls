VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
'
' CLSMDL frmFields  Nov-11-03 17:10
' Prev: Apr-4-03 11:20
' Prev: Nov-20-02  14:45
'
'================================================================================
Dim lngPrevFldID As Long
Dim blnPrevFldRequired As Boolean
Dim strFieldsFormCaller As String, strPrevFldName As String

Private Sub cbxCompID_AfterUpdate()
  Dim frm As Form
  
  With Me
    If Not !cbxCompID.Locked Then
      lngCurrentComp = !cbxCompID
      Call RefreshFilter
      ' If change of co, then synchronize calling form
      If Forms(cstFRM_Main)!cbxSelComp <> lngCurrentComp Then '!cbxCompID Then '
        Set frm = Forms(cstFRM_Main)
        frm!cbxSelComp = lngCurrentComp
        frm!cbxSelComp.Requery
        Set frm = Nothing
      End If
      If strFieldsFormCaller = cstSFRM_Comp Then Call SyncSubRecord("lngCompID", lngCurrentComp)
    End If
  End With
End Sub

Private Sub cbxFldID_AfterUpdate()
  If blnCopyField Then
    Me!cbxFldID.BackColor = vbWhite
    Me!cbxFldID = lngPrevFldID
    DoCmd.RunCommand acCmdSaveRecord
    blnCopyField = False
  End If
  Me!cbxFldID.Requery
  Me!cbxFldID.Locked = True
  Call RefreshFilter
End Sub

Private Sub cbxFldID_BeforeUpdate(Cancel As Integer)
  Dim lngChangeResult As Long
  Dim lngNewFld As Long
  Dim blnNewFldIsDefined As Boolean, blnReset As Boolean, blnDelete As Boolean, blnAbort As Boolean
  Dim strMsg As String, strTitle As String, strNewFldName As String
  strMsg = ""
  
  On Error GoTo cbxFldID_BeforeUpdateErr
  With Me
    lngPrevFldID = Nz(!cbxFldID.OldValue, 1)
    If Not IsNull(Me!cbxFldID) Then
      strPrevFldName = UCase(Me!cbxFldID.Column(1))
      If Not IsNull(Me!cbxFldID.Column(2)) Then blnPrevFldRequired = Me!cbxFldID.Column(2)
    End If
      
    If .NewRecord Then
      If !cbxFldID.Column(3) = "Yes" Then
        strMsg = !cbxFldID.Column(1) & " has already been defined."
        strTitle = "Wrong Choice"
        GoTo Reset
      End If
      
    Else
      ' If event not called on new rec, then it only applies to the Copy function:
      If Not blnCopyField Then GoTo BeforeUpdateExit
      
      blnNewFldIsDefined = False: blnReset = False: blnDelete = False: blnAbort = False
      lngChangeResult = 0
      
      If IsNull(!cbxFldID) Then
        strMsg = "Invalid choice"
        strTitle = "Field Copy Cancelled"
        GoTo Reset
      End If
      
      ' Get new Fld ID & name:
      lngNewFld = !cbxFldID
      strNewFldName = UCase(Me!cbxFldID.Column(1))
        
      'Check if new value is different
      If lngNewFld = lngPrevFldID Then
        strMsg = "You have not selected a different field."
        strTitle = "Field Copy Cancelled"
        GoTo Reset
      End If

      blnNewFldIsDefined = IIf(!cbxFldID.Column(3) = "No", False, True)
       
      If Not blnPrevFldRequired Then  'ask (flag set in cbxFldID_Enter)
        If MsgBox("Do you want to delete the former field after the update?", _
                  vbQuestion + vbYesNoCancel, "Delete Former Field?") = vbYes Then
          blnDelete = True
        ElseIf vbNo Then
          blnDelete = False
        ElseIf vbCancel Then
          strMsg = "User cancelled the operation."
          strTitle = "Field Copy Cancelled"
          GoTo Reset
        End If
      End If
        
      ' Preset the reset flag according to previous answer so that the the values
      ' of that field in the series data table will be automatically reset if the field
      ' is to be deleted from the company fields tbl.
      blnReset = blnDelete
      If Not blnReset Then  ' Only ask if the field is not to be deleted
        If MsgBox("Do you want to reset the former field's values to 0?", _
                  vbQuestion + vbYesNoCancel, "Reset Former Field?") = vbYes Then
          blnReset = True
        ElseIf vbNo Then
          blnReset = False
        ElseIf vbCancel Then
          strMsg = "User cancelled the operation."
          strTitle = "Field Copy Cancelled"
          GoTo Reset
        End If
      End If
            
      ' Display a final message to enable user to cancel the operation: it
      ' is not possible to do so afterwards and the changes cannot be undone.
      strMsg = "Operation(s) requested: " & vbCrLf & _
               "- Copy the amounts of " & strPrevFldName & " into " & strNewFldName & vbCrLf
      If blnDelete Or blnReset Then
        If blnDelete Then
          strMsg = strMsg & "- Delete " & strPrevFldName & vbCrLf
        ElseIf blnReset Then
          strMsg = strMsg & "- Zero-out the amounts of " & strPrevFldName & vbCrLf
        End If
      End If
      strMsg = strMsg & vbCrLf & "Is this what you want to do?"
      strTitle = "Field Copy Confirmation"
      If MsgBox(strMsg, vbQuestion + vbYesNo, strTitle) = vbNo Then Exit Sub
         
      ' Call the function that performs the selected changes: it will hide the form
      ' or close if blnDelete (to prevent record locking).
      Call CopyFieldValues(lngCurrentForecast, lngCurrentComp, _
                           lngPrevFldID, lngNewFld, _
                           blnPrevFldRequired, blnNewFldIsDefined, blnDelete, blnReset, _
                           lngChangeResult)
      If lngChangeResult <> 0 Then  'there was an error
        strMsg = "Error: CopyFieldValues has non-zero result: " & lngChangeResult & vbCrLf
      Else
        strMsg = "The copy operation has completed successfully." & vbCrLf
        If Not blnNewFldIsDefined Then
          strMsg = strMsg & _
                  "Verify that the new field's Accounting Type and Priority" & vbCrLf & _
                  "are valid for this company (defaults were used)."
        End If
      End If
      MsgBox strMsg, vbInformation, "Field Copy Over"
      
    End If
  End With
  
BeforeUpdateExit:
  Exit Sub
  
Reset:
  Me!cbxFldID.Undo
  Cancel = True
  MsgBox strMsg, vbExclamation, strTitle
  GoTo BeforeUpdateExit
  
cbxFldID_BeforeUpdateErr:
  MsgBox "Error (" & Err.Number & ") " & Err.Description, vbExclamation, "cbxFldID_BeforeUpdate"
  Resume BeforeUpdateExit
End Sub

Private Sub cbxFldID_Exit(Cancel As Integer)
  If blnCopyField Then
    beep
    If MsgBox("Select the destination field using the same drop-down box.", vbExclamation + vbOKCancel, _
           "Field Copy Selection") = vbCancel Then
      blnCopyField = False
    Else
      Cancel = True
    End If
  End If
End Sub

Private Sub Form_Activate()
  Me!lblCopy.Visible = (Forms(cstFRM_Main)!sfrmAny.Form.Name = cstSFRM_Series)
  Me!lblCopyHelp.Visible = Me!lblCopy.Visible
End Sub

Private Sub Form_Close()
 Call lblClose_Click
End Sub

Private Sub Form_Current()
  Dim bln As Boolean

  strFieldsFormCaller = Forms(cstFRM_Main)!sfrmAny.Form.Name
  With Me
    bln = .NewRecord
    If bln Then .Cycle = 1 Else .Cycle = 0
  
    If strFieldsFormCaller <> cstSFRM_Series Then
      !lblCopy.ForeColor = cstGrey
      !lblCopyHelp.ForeColor = cstGrey
      !cbxCompID.Locked = False
    Else
      !cbxCompID.Locked = True
      !lblCopy.ForeColor = cstAAColor
      !lblCopyHelp.ForeColor = cstAAColor
    End If
    .AllowAdditions = bln
    !lblNew.Visible = Not bln
    !lblCopy.Visible = Not bln
    !lblCopyHelp.Visible = Not bln
    !lblDelete.Visible = Not bln
    !lblClose.Visible = Not bln
    !lblSave.Visible = bln
    !lblCancel.Visible = bln
    !cbxFldID.Locked = Not bln
    .AllowDeletions = (IsNull(!cbxFldID) Or Not !cbxFldID.Column(2))
  End With
End Sub

Private Sub Form_Open(Cancel As Integer)
  blnFieldsUpdate = False: blnCopyField = False
  
  If Not IsLoaded(cstFRM_Main) Then Cancel = True
  DoCmd.Hourglass True

  With Me
    If Forms(cstFRM_Main)!sfrmAny.Form.Name = cstSFRM_Series Then
      !lblCopy.Visible = True
      !lblCopyHelp.Visible = True
    Else
      !lblCopy.Visible = False
      !lblCopyHelp.Visible = False
    End If
  
    .DataEntry = False
    !cbxCompID = lngCurrentComp
    Call RefreshFilter
  End With
  
  DoCmd.Hourglass False
End Sub

Private Sub lblCancel_Click()
  With Me
    If .Dirty Or .NewRecord Then
      .Undo
      .Cycle = 0
      Call RefreshFilter
    End If
  End With
End Sub

Private Sub lblCancel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblCancel")
End Sub

Private Sub lblCancel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblCancel")
End Sub

Private Sub lblClose_Click()
  blnCopyField = False
  Me.Visible = False
  DoEvents
  DoCmd.Hourglass True
  If strFieldsFormCaller = cstSFRM_Series Then
    If blnFieldsUpdate Then Call RefreshMainGrids
  ElseIf strFieldsFormCaller = cstSFRM_Summary Then
    If blnFieldsUpdate Then Call RefreshSummaryGrid
  End If
  blnFieldsUpdate = False
  DoCmd.Hourglass False
End Sub

Private Sub lblClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblClose")
End Sub

Private Sub lblClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblClose")
End Sub

Private Sub lblCopy_Click()
  Dim strFld As String, strMsg As String
  strFld = "": strMsg = ""
  
  If Me!lblCopy.ForeColor = cstGrey Then
    MsgBox "The field copy function is only available from" & vbCrLf & _
           "the Forecast Series subform as it applies to a" & vbCrLf & _
           "particluar forecast.", vbInformation, "Unavailable Function"
  Else
    Me!cbxFldID.Locked = False
    strFld = UCase(Me!cbxFldID.Column(1))
    strMsg = "Upon your new selection, the amounts of" & vbCrLf & _
              strFld & " will be copied to the new field" & vbCrLf & _
              "across the current forecast series." & vbCrLf & vbCrLf & _
              "Please note:" & vbCrLf & _
              "You will have the option to 'delete' " & strFld & "," & vbCrLf & _
              "or to reset its amounts to 0 if it is required." & vbCrLf & _
              "Do you want to proceed?"
                     
    If MsgBox(strMsg, vbQuestion + vbYesNo, "Copy Field Data?") = vbYes Then
      blnCopyField = True
      Me!cbxFldID.Locked = False
    Else
      blnCopyField = False
      Me!cbxFldID.Locked = True
    End If
    
  End If  ' copy button disabled
  
End Sub

Private Sub lblCopy_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Me!lblCopy.ForeColor = cstGrey Then Exit Sub
  Call ToggleCtlColors(Me, "lblCopy")
End Sub

Private Sub lblCopy_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Me!lblCopy.ForeColor = cstGrey Then Exit Sub
  Call ToggleCtlColors(Me, "lblCopy")
End Sub

Private Sub lblCopyHelp_Click()
  Dim strMsg As String, strLine As String
  Dim iFileNum As Integer
  Const cstFile = "S:\PublicDB\Fundamentals\Portfolio Monitoring\Misc\FieldCopyHelp.txt"
  
  If Me!lblCopyHelp.ForeColor = cstGrey Then Exit Sub
  iFileNum = FreeFile
  Open cstFile For Input As #iFileNum
  Do While Not EOF(iFileNum)
    Line Input #iFileNum, strLine
    strMsg = strMsg & strLine & vbCrLf
  Loop
  Close #iFileNum

  strMsg = strMsg & vbCrLf & _
          "For instance, if you were to click on 'Copy' using the current selection, " & _
          UCase(Me!cbxFldID.Column(1)) & " would be the field to copy data FROM."

  DoCmd.OpenForm cstFRM_Msg, , , , , , strMsg
  
End Sub

Private Sub lblDelete_Click()
  Dim strMsg As String
  strMsg = ""
  
  With Me
    If .AllowDeletions = True Then
      Call DeleteFieldCheck(!cbxFldID, strMsg)
      If Len(strMsg) > 0 Then
        If MsgBox(strMsg, vbQuestion + vbYesNo, "Field Deletion Data Check") = vbNo Then Exit Sub
      End If
      DoCmd.RunCommand acCmdDeleteRecord
      CurrentDb.TableDefs.Refresh
      !cbxFldID.Requery
      .DataEntry = False
      Call RefreshFilter
      blnFieldsUpdate = True
    Else
      MsgBox "You cannot delete the field '" & !cbxFldID.Column(1) & _
             "': it is required for all companies.", vbInformation, "Required Field"
    End If
  End With
End Sub

Private Sub lblDelete_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblDelete")
End Sub

Private Sub lblDelete_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblDelete")
End Sub

Private Sub lblNew_Click()
  On Error GoTo blNew_ClickErr
  With Me
    .DataEntry = True
    .AllowAdditions = True
    DoCmd.GoToRecord , , acNewRec
    !lblNew.Visible = False
  End With

lblNewExit:
  Exit Sub
blNew_ClickErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , "Fields: lblNew_Click"
  Resume lblNewExit
End Sub

Private Sub lblNew_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblNew")
End Sub

Private Sub lblNew_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblNew")
End Sub

Private Sub lblSave_Click()
  DoCmd.Hourglass True
  On Error Resume Next
  With Me
    If IsNull(!lngCompID) Then !lngCompID = !cbxCompID
    .AllowAdditions = True
    DoCmd.RunCommand acCmdSaveRecord
    .DataEntry = False
    
    Call RefreshFilter
    blnFieldsUpdate = True
    
    If Err <> 0 Then
      If Err = 3101 Then
        MsgBox "Only the Comment field is optional: please " & _
                "complete the record or click Cancel", vbExclamation, "Incomplete Record: Not Saved"
        Err.Clear
        Resume Next
      Else
        MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, "frmFields: lblSave_Click"
        DoCmd.Hourglass False
        Exit Sub
      End If
    End If
    .Requery
  End With
  DoCmd.Hourglass False
End Sub

Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblSave")
End Sub

Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblSave")
End Sub

Private Sub lngAcctgTypeID_Change()
  blnFieldsUpdate = True
End Sub

Private Sub lngPriorityID_Change()
  blnFieldsUpdate = True
End Sub

Sub RefreshFilter()
  On Error GoTo RefreshFilterErr
  With Me
    strFieldsFormFilter = "[tblCompFields]![lngCompID]=" & lngCurrentComp
    .Filter = strFieldsFormFilter
    .FilterOn = True
    .Requery
    !cbxFldID.Requery
    !txtTotRecs = "Total: " & .RecordsetClone.RecordCount
  End With

RefreshFilterExit:
  Exit Sub
  
RefreshFilterErr:
  If Err <> 3101 Then 'no matching Priority ID for new fields
    MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, "frmFields:RefreshFilter"
    Resume RefreshFilterExit
  Else
    Err.Clear
    Resume Next
  End If
End Sub

Function DeleteFieldCheck(lngFldIDToDelete As Long, strList As String)
  Dim dbs As DAO.Database
  Dim qdfFieldsSum As DAO.QueryDef
  Dim rstFieldsSum As DAO.Recordset
  Dim strFldName As String, strDesc As String
  Dim f As Integer, iCount As Integer
  Dim blnIncludeForecast As Boolean
  Const cstCoSeriesFieldsSumQry = "qryCoSeriesFieldsSum"  'param = CompID

  strDesc = "": f = 0: iCount = 0:  blnIncludeForecast = False
  On Error GoTo DeleteFieldCheckErr
  '----------------------
  strFldName = DLookup("[txtFldTblName]", "tlkpAllFields", "[lngFldId]=" & lngFldIDToDelete) & ""
  strFldName = Mid$(strFldName, 4) 'remove the type prefix "cur".
  '----------------------
  
  Set dbs = CurrentDb
  Set qdfFieldsSum = dbs.QueryDefs(cstCoSeriesFieldsSumQry)
  qdfFieldsSum.Parameters(0) = lngCurrentComp
  
  Set rstFieldsSum = qdfFieldsSum.OpenRecordset
  With rstFieldsSum
    If .AbsolutePosition <> -1 Then
      .MoveLast
      .MoveFirst
      Do While Not .EOF
      
        For f = 1 To .Fields.Count - 1
          If Not IsNull(.Fields(f)) Then
            If .Fields(f).Value <> 0 Then
            
              If Mid$(.Fields(f).Name, 9) = strFldName Then
                blnIncludeForecast = True
                iCount = iCount + 1
                f = .Fields.Count 'terminate field iteration
              End If
            End If
          End If
        Next f
        
        If blnIncludeForecast Then
          strDesc = strDesc & "   " & .Fields(0) & vbCrLf
          blnIncludeForecast = False
        End If
        
        .MoveNext
      Loop
    End If  '.AbsolutePosition <> -1
    .Close
  End With
  dbs.Close
  
  If Len(strDesc) > 0 Then
    strList = "ATTENTION:" & vbCrLf & "The field you have selected for deletion (" & _
              strFldName & ") has data in the following forecast"
    If iCount > 1 Then
      strList = strList & "s:"
    Else
      strList = strList & ":"
    End If
    strList = strList & vbCrLf & vbCrLf & strDesc & vbCrLf & _
              "If you have not done so, you may want to review the data prior to proceeding." & _
              vbCrLf & vbCrLf & "Do you want to proceed?"
  Else
    strList = "The field '" & strFldName & "' has not been populated in any series." & vbCrLf & _
    "It is OK to delete." & vbCrLf & "Do you want to proceed?"
  End If

  Set rstFieldsSum = Nothing
  Set dbs = Nothing

DeleteFieldCheckExit:
  Exit Function

DeleteFieldCheckErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, "frmFields: DeleteFieldCheck"
  Resume DeleteFieldCheckExit
End Function
