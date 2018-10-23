VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'====================================================================='
' CLSMDL sfrmSeries Dec-08-03 14:35
'   Changed ResetEntryBox; SaveLastEdit
' Nov-11-03 16:35
'     Removed backward tabbing;
'     Refined data entry checks
'
' Prev: Oct-21-03: added window resizing mde
' Prev: Aug-22-02 16:00
'
'=====================================================================
' Form design specs:
'   Form KeyPreview: ON:
'                    Enables selective processing of TAB key so that only this key
'                    will be used to move through the edit grid. This is necessary
'                    due to the editing scheme implemented to edit a non-editable
'                    grid control (grid update via a text box swapping focus with
'                    the edit grid).
'  Controls: NO control is set for TabStop:
'                    This would interfere with the Form's processing all the
'                    KeyDown events.
'  Hidden Controls: - ColNum, RowNum: for debugging purposes: assignments commented out;
'                   - txtColDate, txtColPeriod: Used to update the series:
'                   Along with the forecast id, they are the parameters of a query
'                   that retrieves the original series (i.e. before any changes). They are
'                   populated every time the user doubleclick on a column of the
'                   "display grid" (ocxFlxGridSeries).
'
Public lngClickedCol As Long, lngClickedRow As Long
Dim blnDateRowClicked As Boolean
Dim blnMouseDown As Boolean
Dim blnFirstEntry As Boolean
Const cstOriginY = 405

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' With KeyPreview ON, this sub processes all KeyDown events: only the TAB one is retained.
'
  Dim dte As Long
  If KeyCode <> vbKeyTab Then Exit Sub      ' proceed only if TAB key used
  If lngClickedCol = 0 Then Exit Sub        ' amts grid not being edited
  blnMouseDown = False
  On Error GoTo KeyDownErr
 
  If lngClickedRow = -1 Then
    lngClickedRow = 0 'goto 1st row
  Else
    If lngClickedRow = Me!ocxFlxGridEdit.Rows Then
      lngClickedRow = 0
  '  Else
  '   lngClickedRow = lngClickedRow + 1   ' go to next row
    End If
  End If
  Call SaveLastEdit
  
  lngClickedRow = lngClickedRow + 1
  If (lngClickedRow >= Me!ocxFlxGridEdit.Rows) Then lngClickedRow = 0
  
  With Me
    !ocxFlxGridEdit.Row = lngClickedRow
    !txtNewData.Visible = True
    !txtNewData.SetFocus
  End With
  Call txtNewData_LostFocus
  
  Exit Sub
KeyDownErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, "Form_KeyDown"
End Sub

Private Sub Form_Load()
  Call ResetEntryBox
  Call RefreshMainGrids   'fills grid, sets global var: dteLastColDate
  If CheckDiscontinuity(lngCurrentForecast, 0, False) = True Then
    MsgBox "Click on the 'Check Continuity button for details.", _
            vbExclamation, "Discontinous Forecast"
  End If
  DoCmd.Hourglass False
End Sub

Private Sub lblContinuity_Click()
  If Me!ocxFlxGridEdit.Rows <= 2 Then
    MsgBox "No series displayed!", vbExclamation, "Empty Grid"
    Exit Sub
  End If
  If lngCurrentForecast = 0 Then lngCurrentForecast = Me.Parent!cbxSelForecast
  Call CheckDiscontinuity(lngCurrentForecast, 0, True) 'True=use message form
End Sub

Private Sub lblContinuity_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblContinuity")
End Sub

Private Sub lblContinuity_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblContinuity")
End Sub

Private Sub lblDeleteSeries_Click()
  Dim dte As Date
  Dim m As Integer
  Dim iDeleteErr As Integer
  
  If lngClickedCol = 0 Then
    MsgBox "Doubleclick on the column you want to delete first.", vbExclamation, "Select Series"
    Exit Sub
  End If
  
  dte = CDate(Me!ocxFlxGridSeries.TextMatrix(0, lngClickedCol))
  m = Me!ocxFlxGridSeries.TextMatrix(1, lngClickedCol)
  If MsgBox("Do you want to delete this series time point?" & vbCrLf & vbCrLf & _
            "Forecast: " & Me.Parent!cbxSelForecast.Column(1) & vbCrLf & _
            "Date: " & dte, vbQuestion + vbYesNo, "Time Point Deletion") = vbYes Then
  
    iDeleteErr = DeleteSeries(Me.Parent!cbxSelForecast, dte, m)
    If iDeleteErr <> 0 Then MsgBox "Error in the DeleteSeries function: " & iDeleteErr, vbInformation, "Error"
    
    Forms(cstFRM_Main)!sfrmAny.SourceObject = cstSFRM_Series
  End If
End Sub

Private Sub lblDeleteSeries_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblDeleteSeries")
End Sub

Private Sub lblDeleteSeries_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblDeleteSeries")
End Sub

Private Sub lblExport_Click()
  Call SavePrompt
  DoCmd.Hourglass True
  If Me!ocxFlxGridSeries.Rows < 5 Then
    DoCmd.Hourglass False
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
  Call SavePrompt
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

Private Sub lblNew_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblNew")
End Sub

Private Sub lblNew_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblNew")
End Sub

Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblSave")
End Sub

Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblSave")
End Sub

Private Sub ocxFlxGridEdit_EnterCell() ' Assign grid value to textbox 'OK
  If lngClickedCol = 0 Then Exit Sub  ' The series data has not been copied onto the edit area
  On Error GoTo GridEdit_EnterCellErr
  lngClickedRow = Me!ocxFlxGridEdit.Row
  Me!ocxFlxGridEdit.CellBackColor = Me!ocxFlxGridEdit.BackColorSel
  Me!txtNewData = Me!ocxFlxGridEdit.TextMatrix(lngClickedRow, 0)

  Exit Sub
GridEdit_EnterCellErr:
  If Err.Number <> 381 Then 'subsc out of range (usually after a previous error)
    MsgBox "Error (" & Err.Number & "): " & Err.Description, , "ocxFlxGridEdit_EnterCell"
  End If
End Sub

Private Sub ocxFlxGridEdit_LeaveCell()
  If lngClickedCol = 0 Then Exit Sub
  On Error GoTo ocxFlxGridEdit_LeaveCellErr
  
  If blnFirstEntry Then
    Me!txtNewData = Me!ocxFlxGridEdit.Text
    blnFirstEntry = False
  End If
  Me!ocxFlxGridEdit.CellBackColor = Me!ocxFlxGridEdit.BackColorBkg
  If Not blnFirstEntry Then
    If ((lngClickedRow - 1) = 1) And Not blnMouseDown Then 'validate period
      Select Case CInt(Me!txtNewData) 'Valid periods: 1 , 3 , 6 Or 12
        Case 1, 3, 6, 12 'ok, leave as is
        Case Else
          beep
          Me!txtNewData = Me!txtColPeriod 'set to previous col period (if there is one), or 1.
      End Select
    End If
  End If
  Me!ocxFlxGridEdit.Text = Me!txtNewData

  Exit Sub
ocxFlxGridEdit_LeaveCellErr:
  If Err <> 13 Then
    MsgBox "Error (" & Err.Number & "): " & Err.Description, , "ocxFlxGridEdit_EnterCell"
  End If
End Sub

Private Sub ocxFlxGridEdit_LostFocus()
  If lngClickedCol = 0 Then Exit Sub
  blnFirstEntry = False
  Me!ocxFlxGridEdit.Text = Me!txtNewData
  Me!ocxFlxGridEdit.CellBackColor = Me!ocxFlxGridEdit.BackColorBkg
  With Me!ocxFlxGridSeries ' Unselect row range
    .RowSel = .Row
    .ColSel = .col
  End With
End Sub

Private Sub ocxFlxGridEdit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
  Dim dte As Date
  If lngClickedCol = 0 Then Exit Sub
  blnMouseDown = True
  On Error GoTo ocxFlxGridEdit_MouseDownErr
  
  lngClickedRow = Me!ocxFlxGridEdit.Row
  Me!ocxFlxGridEdit.CellBackColor = Me!ocxFlxGridEdit.BackColorSel
  If lngClickedRow = 0 Then blnDateRowClicked = True 'used in last check before Save: Save_Click
  
  ' Select corresponding row range
  With Me!ocxFlxGridSeries
    ' Row and Col properties must be set before RowSel and ColSel
    .col = 0                ' start selection in this column
    .Row = lngClickedRow    ' start selection in this row
    .ColSel = .Cols - 1       ' end selection in this column
    .RowSel = lngClickedRow ' end selection in this row
  End With
   
  ' Check if 'incoming' date is valid (user will be able to correct other erroneous numeric data
  If lngClickedRow = 0 Then 'date
    dte = GetMonthEndDate(Me!ocxFlxGridEdit.Text)
    ' The function GetMonthEndDate will use current date if input not a date
    Me!ocxFlxGridEdit.Text = Format(dte, cstDateDisp)
  End If
  ' Send cell value to edit box
  Me!txtNewData = Me!ocxFlxGridEdit.Text
  Call MoveTextBox
  
  Exit Sub
ocxFlxGridEdit_MouseDownErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , "ocxFlxGridEdit_MouseDown"
End Sub

Private Sub ocxFlxGridEdit_Scroll()
  Me!ocxFlxGridSeries.TopRow = Me!ocxFlxGridEdit.TopRow + 1
End Sub

Private Sub ocxFlxGridSeries_DblClick()
  Call CopyGridColForEdit ' Copies the clicked column data to the editing area
End Sub

Private Sub ocxFlxGridSeries_Scroll()
  Me!ocxFlxGridEdit.TopRow = Me!ocxFlxGridSeries.TopRow - 1
End Sub

Private Sub lblNew_Click()

  If IsNull(Me.Parent!cbxSelForecast) Then
    MsgBox "This function creates a new financial series for the selected forecast." & vbCrLf & _
          "Please, select a forecast by doubleclicking on a column" & vbCrLf & _
          "to bring the data in the editable grid.", vbExclamation, "New Financial Series Entry Cancelled"
  Else    'open new series form with same forecast id
    Call SavePrompt
    DoCmd.OpenForm cstFRM_NewSer, acNormal, , , acFormAdd, acWindowNormal, Me.Parent!cbxSelForecast
  End If
End Sub

Private Sub lblSave_Click()
  Dim r As Integer, iUpdate As Integer, iMon As Integer, iNewPeriod As Integer
  Dim str As String, strNewDate As String
  Dim dteInitialDte As Date, dteNew As Date
  Dim var() As Variant, varLookup  As Variant
  Dim lngLastEditCol As Long
  On Error GoTo lblSaveErr
  
  If IsNull(Me.Parent!cbxSelForecast) Or lngClickedCol = 0 Then Exit Sub
  
  If Me!ocxFlxGridEdit.Rows <= 2 Then
    MsgBox "No series displayed!", vbExclamation, "Empty Grid"
    Exit Sub
  End If
  
  lngLastEditCol = lngClickedCol  'save locally
  '===========
  Call SaveLastEdit   ': Check value of text box & update edit grid if necessary
    
  Me!txtNewData.Top = cstOriginY  ': needed???
    
  With Me ' Get update function parameters:
    dteInitialDte = CDate(!txtColDate)
    iMon = !txtColPeriod
    If lngR = 0 Then lngR = !ocxFlxGridEdit.Rows
  End With
  
  ' Populate array with the values in the Edit Grid & add the internal Field Names_
  ' in a 2-dim array to pass into the UpdateSeries fctn.
  '
  ReDim var(lngR - 1, lngR - 1)
  For r = 0 To lngR - 1
    ' Process first 2 flds separately (non currency flds, not in tlkpAllFields).
    If r < 2 Then
      If r = 0 Then 'internal date fld name:
        var(0, r) = "dtePeriodEndDate"
        If Not blnDateRowClicked Then    'process date: since it was not clicked, no check performed
          dteNew = GetMonthEndDate(Me!ocxFlxGridEdit.TextMatrix(r, 0))
          Me!ocxFlxGridEdit.TextMatrix(r, 0) = Format(dteNew, cstDateDisp)
        End If
        var(1, r) = CLng(CDate(Me!ocxFlxGridEdit.TextMatrix(r, 0)))
        
      ElseIf r = 1 Then
        var(0, r) = "intMonths"
        var(1, r) = CLng(Me!ocxFlxGridEdit.TextMatrix(r, 0))
      End If
        
    Else
      str = Me!ocxFlxGridSeries.TextMatrix(r, 0) 'save disp name for lookup
      varLookup = DLookup("[txtFldTblName]", "tlkpAllFields", "[txtDispName]='" & str & "'")
      If IsNull(varLookup) Then
        MsgBox "The internal field name lookup failed to find a match for " & _
                str & " in tlkpAllFields." & vbCrLf & _
                "Field preceding failure: " & var(0, r - 1), vbCritical, "Field Name Mismatch"
        Call Form_Load
        Exit Sub
      Else
        var(0, r) = varLookup
        var(1, r) = Me!ocxFlxGridEdit.TextMatrix(r, 0)
      End If
    End If
  Next r
  
  ' Call to Update function:
  If lngCurrentForecast = 0 Then lngCurrentForecast = Me.Parent!cbxSelForecast
  
  iUpdate = UpdateSeries(dteInitialDte, iMon, var)  'uses current forecast
  '
  If iUpdate <> 0 Then
    MsgBox "An error occurred in the function 'UpdateSeries': " & iUpdate, _
           vbExclamation, "Error"
    Call Form_Load 'call continuity check
    Exit Sub
  End If

  Call Form_Load  'dteLastColDate set by RefreshMainGrids
  Me!ocxFlxGridSeries.LeftCol = lngLastEditCol
  'If the last col was updated, then enter next consecutive col?
  If dteInitialDte >= dteLastColDate Then 'dteLastColDate set by RefreshMainGrids called in Form_Load
    If MsgBox("Enter next point?", vbQuestion + vbYesNo, "Continue data entry?") = vbYes Then
      DoCmd.OpenForm cstFRM_NewSer, acNormal, , , acFormAdd, acHidden, Me.Parent!cbxSelForecast
      Call SaveNewSeries
    End If
  End If
  
  Exit Sub
lblSaveErr:
  If Err = 2165 Then
    Err.Clear
    Resume
  Else
    MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, "lblSave_click"
  End If
End Sub

Private Sub txtNewData_GotFocus()
  If lngClickedCol = 0 Then Exit Sub
  With Me!txtNewData                      ' Highlight content
    .SelStart = 0                         ' Start selection at beginning.
    .SelLength = Len(.Text)               ' Length of text in txtNewData.
    .Value = Me!ocxFlxGridEdit.TextMatrix(lngClickedRow, 0)
  End With
  Call MoveTextBox
End Sub

Private Sub txtNewData_LostFocus()
  Dim dte As Date
  Dim iPeriod As Integer
  Dim blnEntryErr As Boolean, blnIsADate As Boolean, blnIsANum As Boolean, blnIsAPeriod As Boolean
  Dim varPrevData As Variant
  
  blnEntryErr = False: blnIsADate = False: blnIsANum = False: blnIsAPeriod = False:
  If lngClickedCol = 0 Then Exit Sub
  On Error GoTo txtNewData_LostFocusErr
  
  With Me
    ' Process date cell separately for formating:
    If lngClickedRow = 0 Then    'Date row
    
      blnIsADate = IsDate(!txtNewData.Text)
      If Not blnIsADate Then
        ' The data entered is not a date: copy data from original grid
        blnEntryErr = True
        beep
        !txtNewData.Undo
        dte = GetMonthEndDate(!ocxFlxGridSeries.TextMatrix(lngClickedRow, lngClickedCol))
      Else  ' Get last day of month:
        dte = GetMonthEndDate(!txtNewData)
      End If
      
      ' Format date field & update the textbox
      If Not blnEntryErr Then  ' no field has been reset
        varPrevData = !txtNewData
      Else
        varPrevData = Format(dte, cstDateDisp)
      End If
      
    Else ' row > 0 for all other cells:
    
      blnIsANum = IsNumeric(!txtNewData)
      blnIsADate = IsDate(!txtNewData.Text)
      
      ' Process periodicity cell (number of months), integer format, not currency:
      If lngClickedRow = 1 Then
        blnIsAPeriod = (!txtNewData = 1 Or !txtNewData = 3 Or !txtNewData = 6 Or !txtNewData = 12)
        If Not blnIsAPeriod Then
          blnEntryErr = True
          beep
          !txtNewData.Undo      'reset
          varPrevData = CInt(!txtColPeriod) '(set to 1 or the prev col period if there is one)
        End If
      Else  ' row >1
        If Not blnIsANum Or blnIsADate Then      '
          blnEntryErr = True
          beep
          !txtNewData.Undo      'reset
          ' With the numerical fields, re-enter the original values:
          varPrevData = Format(CCur(!ocxFlxGridSeries.TextMatrix(lngClickedRow - 1, lngClickedCol)), cstCurrDisp)
        End If ' Not blnIsANum Or blnIsADate
      
      End If ' click row > 1
    
      If Not blnEntryErr Then  ' no field has been reset
        If lngClickedRow = 1 Then
          varPrevData = CInt(!txtNewData)
        Else
          varPrevData = Format(CCur(Nz(!txtNewData, 0)), cstCurrDisp)
        End If
      End If
    
    End If
      
    ' Update box & grid cell with the validated data:
    !txtNewData = varPrevData
    !ocxFlxGridEdit.Text = !txtNewData
  End With
 
  Call MoveTextBox

  Exit Sub
txtNewData_LostFocusErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, "txtNewData_LostFocus"
End Sub

'=====================================================================
Sub MoveTextBox()        ' Moves Textbox next to active cell.
  With Me!txtNewData
    .Visible = True
    .SetFocus
    .Top = Me!ocxFlxGridEdit.Top + Me!ocxFlxGridEdit.CellTop
  End With
  ' Highlight Series Table Row:
  If lngClickedRow = Me!ocxFlxGridEdit.Rows Then lngClickedRow = 0
  With Me!ocxFlxGridSeries
    ' Row and Col properties must be set before RowSel and ColSel
    .col = 0                    ' start selection in this column
    .Row = lngClickedRow        ' start selection in this row
    .ColSel = .Cols - 1         ' end selection in this column: hilight entire row
    .RowSel = lngClickedRow     ' end selection in this row
  End With
End Sub

Sub ResetEntryBox()
  With Me
    !txtNewData.Top = cstOriginY
    !txtNewData = "<end date>"
    !txtNewData.Visible = False
    !txtColDate = Null            ' Hidden ctl used for updating series
    !txtColPeriod = Null          ' Hidden ctl used for updating series
    blnFirstEntry = True
    blnDateRowClicked = False
    lngClickedCol = 0
    lngClickedRow = -1 'To find out if data was entered
  End With
End Sub

Sub SaveLastEdit()
  'Called in case TAB was not pressed or edit grid not clicked when Save called
  Dim iPeriod As Integer
  Dim dte As Long
  Dim blnOK As Boolean
  blnOK = False
  On Error GoTo SaveLastEditErr

  With Me
    !txtNewData.Visible = True
    !txtNewData.SetFocus
    If IsNumeric(!txtNewData.Text) Then
      If lngClickedRow = 1 Then  'validate period before leaving control on forward
        iPeriod = CInt(!txtNewData) 'Valid periods: 1 , 3 , 6 Or 12
        Select Case iPeriod
          Case 1, 3, 6, 12
          Case Else
            beep
            iPeriod = !txtColPeriod 'set to previous col period (if there is one), or 1.
        End Select
        !ocxFlxGridEdit.TextMatrix(1, 0) = iPeriod
      End If
      blnOK = True
    ElseIf IsDate(!txtNewData.Text) Then
      If lngClickedRow = 0 Then
        dte = GetMonthEndDate(!txtNewData.Text)
        If dte <> CDate(!ocxFlxGridEdit.TextMatrix(0, 0)) Then
          !txtNewData.Text = Format(dte, cstDateDisp)
          blnOK = True
        End If
      Else
        beep
      End If
    End If 'numeric or date

    If blnOK Then !ocxFlxGridEdit.SetFocus
  End With

  Exit Sub
SaveLastEditErr:
  Err.Clear
End Sub

Sub SavePrompt()
  If lngClickedCol = 0 Or lngClickedRow = -1 Then Exit Sub
  If MsgBox("Do you want to save your changes?", vbQuestion + vbYesNo, "Update Period?") = vbYes Then
    Call lblSave_Click
  End If
End Sub

Sub CopyGridColForEdit()
Dim r As Integer, iPeriod As Integer
  
  Call SavePrompt
  Call ResetEntryBox
  
  lngClickedCol = Me!ocxFlxGridSeries.col
  With Me
    !ocxFlxGridSeries.ColSel = lngClickedCol ' start selection in this column
    !ocxFlxGridSeries.RowSel = !ocxFlxGridSeries.Rows - 1     ' end selection in this row
    
    'Copy col data to other grid:
    For r = 0 To !ocxFlxGridSeries.RowSel
      If r < 2 Then   'Do not format
       ' Validate periodicity from tbl data:
        If r = 1 Then
          iPeriod = CInt(!ocxFlxGridSeries.TextMatrix(1, lngClickedCol))  'Valid periods: 1 , 3 , 6 Or 12
          Select Case iPeriod
            Case 1, 3, 6, 12
            Case Else
              If (lngClickedCol - 1) > 0 Then
                iPeriod = !ocxFlxGridSeries.TextMatrix(1, lngClickedCol - 1)
              Else
                iPeriod = 1       'reset to 1
              End If
          End Select
          !ocxFlxGridEdit.TextMatrix(r, 0) = iPeriod
        Else 'r=0: date
          !ocxFlxGridEdit.TextMatrix(r, 0) = !ocxFlxGridSeries.TextMatrix(r, lngClickedCol)
        End If
      Else ' curr amounts cells
        !ocxFlxGridEdit.TextMatrix(r, 0) = Format(!ocxFlxGridSeries.TextMatrix(r, lngClickedCol), cstCurrDisp)
      End If
    Next r
    
    ' Populate other ctls:
    !txtColDate = !ocxFlxGridEdit.TextMatrix(0, 0)
    !txtColPeriod = iPeriod
    !ocxFlxGridSeries.RowSel = 0     ': Clear highlight
  End With
End Sub
