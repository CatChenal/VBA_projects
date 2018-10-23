Attribute VB_Name = "MDL Calendar"
Option Compare Database
Option Explicit
'================================================================================
' DBFrontEnd
' MDL Calendar May-9-02 12:20
'
'================================================================================
Const cstMDL = "Calendar"
Public Const cstFRM_Cal = "frmCalendar"

Function UpdateDateFromCal(Optional blnUndo As Boolean)
  Dim frmCal As Form
  Dim frmUpd As Form
  Dim ctl As Control
  Dim strForm As String, strControl As String
  
  On Error GoTo DateFromCalErr
  Set frmCal = Forms(cstFRM_Cal)
  With frmCal
    If (Len(!txtCtl) = 0 Or Len(!txtForm) = 0) Then
      GoTo DateFromCalExit  'both are needed
    End If
    strControl = !txtCtl
    strForm = !txtForm
    Set frmUpd = Forms(strForm)
    
    Set ctl = frmUpd.Controls(strControl)
    If IsMissing(blnUndo) Or Not blnUndo Then
       ctl = !acxCal.Value
    Else
      If blnUndo Then
        If CDate(!PreVal) <> Date Then
          ctl = Format(!PreVal, "Medium Date")
        Else
          ctl.Value = ctl.OldValue
          'ctl.Undo
        End If
      End If
    End If
  End With
  
DateFromCalExit:
  Set ctl = Nothing
  Set frmUpd = Nothing
  Set frmCal = Nothing
  Exit Function

DateFromCalErr:
  MsgBox "Error: (" & Err & ") " & Err.Description, vbExclamation, cstMDL & ": UpdateDateFromCal"
  Resume DateFromCalExit
End Function

Function GetMonthEndDate(varDateInput As Variant) As Date
  If Not IsDate(varDateInput) Then  'use current date
    GetMonthEndDate = DateSerial(Year(Date), Month(Date) + 1, 0)
  Else
    GetMonthEndDate = DateSerial(Year(varDateInput), Month(varDateInput) + 1, 0)
  End If
End Function

Public Sub OpenCalendar(frmCaller As Form, strCallCtl As String)
  Dim dteOpenDate  As Date
  Dim frmCal As Form
  Dim txt As TextBox
  Dim str$
  On Error GoTo OpenCalErr
  
  str$ = frmCaller.Name
  Set txt = frmCaller.Controls(strCallCtl)
  If IsNull(txt.Value) Or Not IsDate(txt.Value) Then
    dteOpenDate = Date
  Else
    dteOpenDate = CDate(txt.Value)
  End If

  DoCmd.OpenForm cstFRM_Cal, , , , , acHidden, dteOpenDate
  Set frmCal = Forms(cstFRM_Cal)
  
  'lngMoveRight = (Screen.ActiveForm.WindowWidth / 2) - (frmCal.WindowWidth / 2)
  'DoCmd.MoveSize Right:=lngMoveRight, Down:=cstMoveDown

  frmCal.Visible = True
  With frmCal
    !acxCal = dteOpenDate
    !txtForm = str$
    !txtCtl = strCallCtl
  End With
  
OpenCalExit:
  Set txt = Nothing
  Set frmCal = Nothing
  Set frmCaller = Nothing
  Exit Sub

OpenCalErr:
  MsgBox "Error: (" & Err & ") " & Err.Description & vbCrLf & _
         "Calling Form: " & str$ & vbCrLf & _
         "Calling Control: " & strCallCtl, vbExclamation, cstMDL & ":OpenCalendar"
  Resume OpenCalExit
End Sub

