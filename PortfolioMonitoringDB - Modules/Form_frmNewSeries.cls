VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmNewSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
'
' CLSMDL frmNewSeries Jun-7-02 15:10
'
'================================================================================
Dim sfrm As Form
  
Private Sub cmbCalendar_Click()
  Call OpenCalendar(Me, "dtePeriodEndDate")
End Sub

Private Sub Form_Close()
  DoCmd.Hourglass True
  Set sfrm = Nothing
  Call RefreshMainGrids
  DoCmd.Hourglass False
End Sub

Private Sub Form_Current()
  With Me
    !dtePeriodEndDate = RefreshEndDate
    !cbxForecastID = Forms(cstFRM_Main)!cbxSelForecast
  End With
End Sub

Private Sub Form_Load()
  Dim iMon As Integer
  
  ' Set default val for opg from value on grid:
  Set sfrm = Forms(cstFRM_Main)!sfrmAny.Form
  lngC = sfrm!ocxFlxGridSeries.Cols - 1
  If Len(sfrm!ocxFlxGridSeries.TextMatrix(1, lngC)) > 0 Then
    iMon = sfrm!ocxFlxGridSeries.TextMatrix(1, lngC)
  End If
  Set sfrm = Nothing
  
  With Me
    !cbxForecastID.SetFocus
    !cbxForecastID = Forms(cstFRM_Main)!cbxSelForecast.ItemData(0)
    !opgPeriod.SetFocus
    Select Case iMon
      Case 3, 6, 12
        !opgPeriod = iMon
      Case Else
        !opgPeriod = 1
    End Select
  End With
    
End Sub

Private Sub Form_Open(Cancel As Integer)
  Dim iMon As Integer
  Dim strMsg As String
  iMon = 0: strMsg = ""
  
  On Error GoTo Form_OpenCancel
  If Len(Me.OpenArgs & "") = 0 Then
    strMsg = "This form is used to add a new financial series to a company forecast." & vbCrLf
    GoTo Form_OpenCancel
  End If
  
  If IsNull(Me!cbxForecastID) Then
    strMsg = "The selection box should display the current forecast name." & vbCrLf
    GoTo Form_OpenCancel
  End If
  
Form_OpenExit:
  Exit Sub
  
Form_OpenCancel:
  Cancel = True
  If Err = 0 Then
    strMsg = strMsg & "You need to open the Main form and go to the Forecast Data page " & _
                    "before being able to do so."
  Else
    strMsg = "Error (" & Err.Number & "): " & Err.Description
  End If
  MsgBox strMsg, vbExclamation, "Cancelling New Financial Series Entry"
  Resume Form_OpenExit
End Sub

Private Sub lblCancel_Click()
 If Me.Dirty Then
    Me.Undo
    Set sfrm = Nothing
    DoCmd.Close acForm, cstFRM_NewSer
  End If
End Sub

Private Sub lblClose_Click() ' Save caption
  Dim blnSave As Boolean
  
  If Me.Dirty Then
    If MsgBox("Save this new series?", vbQuestion + vbYesNo, "Save New Entry?") = vbYes Then
      Me!intMonths = Me!opgPeriod.Value
      DoCmd.RunCommand acCmdSaveRecord
    Else
      Call lblCancel_Click
    End If
  End If
  DoCmd.Close acForm, cstFRM_NewSer
End Sub

Private Sub opgPeriod_AfterUpdate()
  Me!intMonths = Me!opgPeriod.Value  'assign to tbl variable
  Me!dtePeriodEndDate = RefreshEndDate
End Sub

Function RefreshEndDate() As Date
  Dim dteTemp As Date, dteLast As Date
  On Error GoTo RefreshEndDateErr
  
  Set sfrm = Forms(cstFRM_Main)!sfrmAny.Form
  If lngC = 0 Then lngC = sfrm!ocxFlxGridSeries.Cols - 1

  ' Get last date in grid:
  If Not IsDate(sfrm!ocxFlxGridSeries.TextMatrix(0, lngC)) Then
    dteTemp = Date 'reset
  Else
    dteTemp = CDate(sfrm!ocxFlxGridSeries.TextMatrix(0, lngC))
  End If
  Set sfrm = Nothing

  ' Get the next month date according to given period:
  dteLast = DateSerial(Year(dteTemp), Month(dteTemp) + Me!opgPeriod, 1)
  ' Get the end of that month:
  RefreshEndDate = GetMonthEndDate(dteLast)
  
RefreshEndDateExit:
  Set sfrm = Nothing
  Exit Function

RefreshEndDateErr:
  RefreshEndDate = Date
  MsgBox "RefreshEndDateErr (" & Err.Number & "): " & Err.Description
  Resume RefreshEndDateExit
End Function
