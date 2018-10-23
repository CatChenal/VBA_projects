VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
'
' CLSMDL sfrmAnalysis: Apr-13-04
' Prev: Mar-15-04
'
'================================================================================

Private Sub cbxSelYear_Change()
  DoCmd.Hourglass True
  Call RefreshQEndDate(False)
  Call FillFileTree(Me!cbxSelYear, Me!opgQ)
  DoCmd.Hourglass False
End Sub

Private Sub Form_Deactivate()
   Forms(cstFRM_Main)!dteQEndDate = Date
End Sub

Private Sub Form_Load()
  DoCmd.Hourglass True
  iPrevNodeIdx = 0
  Call RefreshQEndDate(False)
  Call FillFileTree(Year(Forms(cstFRM_Main)!dteQEndDate), Me!opgQ)
  DoCmd.Hourglass False
End Sub

Private Sub lblAddReview_Click()
' Label acts as a button with 2 options depending on caption: Add or Print
  Const cstFile = cstFrontEndFolder & "List of All Quarterly Reviews.txt"
  Dim strOutput As String
  Dim lngOutcome As Long
  On Error GoTo ProcErr
  
  If InStr(Me!lblAddReview.Caption, "Add") > 0 Then 'add review
    If AddReviewToFolder <> 0 Then Exit Sub
    Call FillFileTree(Me!cbxSelYear, Me!opgQ)
  Else
    Call GetDocList(strOutput)
    strOutput = ".LOG" & vbCrLf & strOutput
    Call SaveToFile(cstFile, strOutput, True) 'Display w/Notepad:
    lngOutcome = Shell(cstWinAppPath & "NOTEPAD.EXE " & cstFile, vbNormalFocus)
  End If
  
  Exit Sub
  
ProcErr:
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "lblAddReview_Click"
End Sub

Private Sub lblCompile_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblCompile")
End Sub

Private Sub lblCompile_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblCompile")
End Sub

Private Sub lblCompile_Click()
  DoCmd.Hourglass True
  Call CompileAllDocs
  DoCmd.Hourglass False
End Sub

Private Sub lblCompute_Click()
  If NoDataCheck Then Exit Sub
  DoCmd.Hourglass True
  lngCurrentForecast = Me.Parent!cbxSelForecast
  Call CalcFinInfo(lngCurrentForecast, Me!cbxSelPeriodFrom, Me!cbxSelPeriodTo)
  DoCmd.Hourglass False
End Sub

Private Sub lblCompute_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblCompute")
End Sub

Private Sub lblCompute_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblCompute")
End Sub

Private Sub lblMissingReviewRpt_Click()
 Call FillFileTree(Me!cbxSelYear, Me!opgQ, True)
End Sub

Private Sub lblMissingReviewRpt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblMissingReviewRpt")
End Sub

Private Sub lblMissingReviewRpt_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblMissingReviewRpt")
End Sub

Private Sub lblPrint_Click()
  Dim appXL As Excel.Application
  Dim wbk As Excel.Workbook
  Dim str As String
  On Error GoTo lblPrint_ClickErr
  
  If NoDataCheck Then Exit Sub
  If Not Me!oleuOutput.Visible Then
    MsgBox "Click on 'Compute' first.", vbExclamation, "Nothing to Print"
    Exit Sub
  End If
  DoCmd.Hourglass True
  str = Me!oleuOutput.SourceDoc
  Debug.Print str
  Set appXL = CreateObject("Excel.Application")
  appXL.WindowState = xlNormal
  Set wbk = appXL.Workbooks.Open(str, , , , , , , , , False, , , False)
  appXL.Visible = True
  wbk.Sheets(1).PageSetup.Zoom = 100
  wbk.PrintPreview False
  While Left(appXL.StatusBar, 7) = "Preview"
  Wend
  appXL.Visible = False
  appXL.Quit
  
lblPrint_ClickExit:
  DoCmd.Hourglass False
  Set wbk = Nothing
  Set appXL = Nothing
  Exit Sub
  
lblPrint_ClickErr:
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "lblPrint_Click"
  Resume lblPrint_ClickExit
End Sub

Private Sub lblPrint_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblPrint")
End Sub

Private Sub lblPrint_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call ToggleCtlColors(Me, "lblPrint")
End Sub

Private Sub ocxTree_BeforeLabelEdit(Cancel As Integer)
  Cancel = True 'to prevent edits bu user
End Sub

Private Sub ocxTree_Click()
  Call TreeClick
End Sub

Private Sub ocxTree_DblClick()
  Dim nod As Node
  Dim strFile As String
  On Error GoTo ocxTree_DblClickErr
  If Me!ocxTree.SelectedItem.Index = 1 Then Exit Sub 'not applicable on root node
  
  DoCmd.Hourglass True
  Set nod = Me!ocxTree.SelectedItem
  strFile = nod.Key

  If InStr(strFile, "\") = 0 Then 'comp hdr
    If nod.Children = 0 Then
      GoTo ocxTree_DblClickExit
    Else
      nod.Expanded = True
      If nod.Children = 1 Then
        strFile = nod.Child.Key
      Else
        MsgBox "Please select which file to open.", vbExclamation, "No File Selected"
        GoTo ocxTree_DblClickExit
      End If
    End If
  End If
  
  ' Use Read Only label visible prop to open doc accordingly:
  Call OpenPFMReview(strFile, Me!lblReadOnly.Visible)
  
ocxTree_DblClickExit:
  DoCmd.Hourglass False
  Set nod = Nothing
  Exit Sub
  
ocxTree_DblClickErr:
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "ocxTree_DblClick"
  Resume ocxTree_DblClickExit
End Sub

Private Sub ocxTree_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                              ByVal x As Long, ByVal y As Long)
Dim str As String
  If (Button And acRightButton) Then
    Call ocxTree_DblClick   ' Open doc on rightclicking
  ElseIf (Button And acLeftButton) Then  'display path in tip text
    Me!ocxTree.ControlTipText = Me!ocxTree.SelectedItem.Key
  End If
End Sub

Private Sub ocxTree_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                            ByVal x As Long, ByVal y As Long)
  If (Button And acLeftButton) Then  'restore tree obj ControlTipText
    Me!ocxTree.ControlTipText = Me!ocxTree.Tag
  End If
End Sub

Private Sub opgQ_Click()
  DoCmd.Hourglass True
  Call RefreshQEndDate(Me!optAll)
  Call FillFileTree(Me!cbxSelYear, Me!opgQ)
  DoCmd.Hourglass False
End Sub

Private Sub optAll_Click()
  Call FillFileTree(Me!cbxSelYear, Me!opgQ)
End Sub

Private Sub cbxSelPeriodFrom_AfterUpdate()
  Me!oleuOutput.Visible = False
  Me!lblTitlePeriod.Caption = Me!cbxSelPeriodFrom
  Me!cbxSelPeriodTo.Requery
  Me!cbxSelPeriodTo = Null
End Sub

Private Sub cbxSelPeriodTo_AfterUpdate()
  If Me!cbxSelPeriodTo < Me!cbxSelPeriodFrom Then
    beep
    Me!cbxSelPeriodTo = Me!cbxSelPeriodFrom
  End If
  Me!oleuOutput.Visible = False
  Me!lblTitlePeriod.Caption = Me!cbxSelPeriodFrom & " To " & Me!cbxSelPeriodTo
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

Function NoDataCheck() As Integer '-1= no go
  blnNoFrc = (Me.Parent!cbxSelForecast.ListCount = 0)
  blnNoPoints = (Me!cbxSelPeriodFrom.ListCount = 0)
  If blnNoFrc Or blnNoPoints Or Me!optAll Then NoDataCheck = True
  If blnNoFrc Then
    MsgBox "There are no forecasts for this company", vbExclamation, "No Financial Info"
  ElseIf blnNoPoints Then
    MsgBox "There are no data points for this period.", vbExclamation, "No Financial Info"
  ElseIf Me!optAll Then
    MsgBox "Select a single period.", vbExclamation, "All Periods Selected"
  End If
End Function

Sub RefreshQEndDate(blnAllPeriods As Boolean)
  If blnAllPeriods Then
    Me.Parent!dteQEndDate = Date
  Else
    Me.Parent!dteQEndDate = GetGivenQtrDate(Me!opgQ, Me!cbxSelYear, 0)
  End If
End Sub
