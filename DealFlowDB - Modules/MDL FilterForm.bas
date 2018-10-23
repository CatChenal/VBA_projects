Attribute VB_Name = "MDL FilterForm"
Option Compare Database
Option Explicit
'===============================================================================
' DBFrontEnd
' MDL FilterForm  Sep-05-03 14:40
'   Changes: GetFilterName: modified desciptive.
'   Prev: Sept-03-02 14:35
'   Changes: Added fctns: UndoLastEdits & AddToEditList to clear ctls
'                         when Cancel is clicked.
'
'   Prev: Jul-25-02
'   Changes in: GetFilteredRecords proc to requery main form.
'             : ApplyFormFilter proc: simplified caption change.
'   Prev: Jun-14-02 16:00
'
'===============================================================================
Const cstMDL = "MDL FilterForm"
Public Const cstDefWidth = 4245     ' InsideWidth: 2.95"
Public Const cstDefHeight = 6990    ' InsideHeight: 4.85"

Sub ResetFilterWindow(iDefaultWidth As Integer, iDefaultHeight As Integer)
' If iDefaultWidth or iDefaultHeight =0 then form window is fitted to form size.
  Dim frm As Form
  Set frm = Forms(cstFilterForm)
  If iDefaultWidth = 0 Or iDefaultHeight = 0 Then
    Call ResetWindowSize(frm)
  Else
    frm.InsideWidth = iDefaultWidth
    frm.InsideHeight = iDefaultHeight
  End If
  Set frm = Nothing
End Sub

Sub ResetWindowSize(frm As Form)
  Dim intTotalFormHeight As Integer, intTotalFormWidth As Integer
  With frm
    intTotalFormHeight = .Section(acHeader).Height + _
                         .Section(acDetail).Height + .Section(acFooter).Height
    intTotalFormWidth = .Width
    If .InsideWidth <> intTotalFormWidth Then .InsideWidth = intTotalFormWidth
    If .InsideHeight <> intTotalFormHeight Then .InsideHeight = intTotalFormHeight
  End With
End Sub

Function FormFilterIsSet() As Boolean
  Dim frm As Form
  Dim ctl As Control
  Dim bln As Boolean
  Dim strFilterText As String
  Dim c As Integer, iPrevIdx As Integer, iIdx As Integer
  Dim lngTxt As Long
  Dim varCtlNames As Variant
  '--------------------------------
  iPrevIdx = -1
  '--------------------------------
  On Error GoTo ProcErr
    
  Set frm = Forms(cstFilterForm)
  frm!tbxFilter = strFilterText
  varCtlNames = Array("cbxSelYear1", "cbxSelYear2", "cbxSelMonth1", "cbxSelMonth2", _
                      "cbxSelQtr1", "cbxSelQtr2", "tbxSelDate1", "tbxSelDate2", _
                      "cbxSelAnalyst", "cbxSelDisposition", "cbxSelSource", "cbxSelSourceType", _
                      "cbxSelSponsor", "cbxSelIssuer", "cbxSelControl", "cbxSelIndustry")
                    
  For c = 0 To UBound(varCtlNames)
    Set ctl = frm.Controls(varCtlNames(c))
    iIdx = ctl.TabIndex
      
    If ctl.Enabled = True Then
      If Not IsNull(ctl) Then
        bln = True
        
        If c < 8 Then 'date criteria: different formatting
          If iIdx = iPrevIdx + 1 Then 'second of a pair (range)
          
            If ctl.ControlType = acTextBox Then
              strFilterText = strFilterText & String(3, " ") & ctl.Tag & " " & ctl '& vbCrLf
            ElseIf ctl.ControlType = acComboBox Then
              ctl.SetFocus
              strFilterText = strFilterText & String(3, " ") & ctl.Tag & " " & ctl.Text   '& vbCrLf
            End If
            
          Else 'new date range ctl
            If ctl.ControlType = acTextBox Then
              strFilterText = strFilterText & vbCrLf & ctl.Tag & " " & ctl
            ElseIf ctl.ControlType = acComboBox Then
              ctl.SetFocus
              strFilterText = strFilterText & vbCrLf & ctl.Tag & " " & ctl.Text
            End If
          End If ' idx is second of a pair
          iPrevIdx = iIdx
       
        Else    ' c >=8: all other cbx
          ctl.SetFocus
          strFilterText = strFilterText & vbCrLf & ctl.Tag & " " & ctl.Text
        End If
      End If 'not is null
    End If 'enabled
    Set ctl = Nothing
  Next c
  lngTxt = Len(strFilterText)
  If lngTxt >= 2 Then
    strFilterText = Right(strFilterText, lngTxt - 2)    'remove first CR + LF
    'Debug.Print "strFilterText: " & strFilterText
  End If
  frm!tbxFilter = strFilterText

ProcExit:
  Set frm = Nothing
  FormFilterIsSet = bln
  Exit Function
  
ProcErr:
  bln = False
  MsgBox "Error: " & Err & "; " & Err.Description, vbExclamation, "FormFilterIsSet"
  Resume ProcExit
End Function

Sub ShowFilterForm(strCallerForm As String, blnAdvanced As Boolean)
  Dim frm As Form
  On Error Resume Next
  
  Set frm = Forms(cstFilterForm)
  frm!tbxCallingForm = strCallerForm
  
  If blnAdvanced Then
    Call ResetFilterWindow(1000, 1000) 'to be determined when
    'ctls for Fin and Sec data are added to the form
  Else
    Call ResetFilterWindow(cstDefWidth, cstDefHeight)
  End If
  
  frm.Visible = True
  If Err <> 0 Then
    If Err = 2450 Then '"Can't find the form"
      DoCmd.RunMacro "AutoExec"
      Err.Clear
      frm.Visible = True
    Else
      MsgBox "Error: " & Err & "; " & Err.Description, vbExclamation, "ShowFilterForm"
    End If
  End If
End Sub

Sub ResetFilterForm()
  With Forms(cstFilterForm)
    !chkMonth = 0
    Call DateSelectionCheckboxClick(2)
    !chkQuarter = 0
    Call DateSelectionCheckboxClick(3)
    !chkRange = 0
    Call DateSelectionCheckboxClick(4)
    !chkYear = 0
    Call DateSelectionCheckboxClick(1)
  
    ' Other fields:
    !cbxSelAnalyst = Null
    !cbxSelDisposition = Null
    !cbxSelSource = Null
    !cbxSelSourceType = Null
    !cbxSelSponsor = Null
    !cbxSelIssuer = Null
    !cbxSelIndustry = Null
    !cbxSelControl = Null
  End With
End Sub

Function DateSelectionCheckboxClick(intCheckboxTag As Integer)
  With Forms(cstFilterForm)
  
    If intCheckboxTag <> 4 Then 'other than date range
      ' Set opt 4 to 0 & disable its related ctls
      Call NoDateRangeCheckboxClick
      Select Case intCheckboxTag
        Case 1
          Call YearRangeCheckboxClick
        Case 2
          Call MonthRangeCheckboxClick
        Case 3
          Call QtrRangeCheckboxClick
        Case Else
          Exit Function
      End Select
    Else  'date range clicked
      ' Set 1, 2, & 3 to 0 & disable their related ctls
      !tbxSelDate1.Enabled = !chkRange
      !tbxSelDate2.Enabled = !chkRange
      
      !chkYear = 0
      If !cbxSelYear1.Enabled Then
        !cbxSelYear1 = Null: !cbxSelYear2 = Null
        !cbxSelYear1.Enabled = False: !cbxSelYear2.Enabled = False
      End If
      
      !chkMonth = 0
      If !cbxSelMonth1.Enabled Then
        !cbxSelMonth1 = Null: !cbxSelMonth2 = Null
        !cbxSelMonth1.Enabled = False: !cbxSelMonth2.Enabled = False
      End If
      
      !chkQuarter = 0
      If !cbxSelQtr1.Enabled Then
        !cbxSelQtr1 = Null: !cbxSelQtr2 = Null
        !cbxSelQtr1.Enabled = False: !cbxSelQtr2.Enabled = False
      End If
    End If
  End With
  
End Function

Sub NoDateRangeCheckboxClick()
' Reset & disable Date range ctls
  With Forms(cstFilterForm)
    !chkRange = 0
    If !tbxSelDate1.Enabled Then
      !tbxSelDate1 = Null: !tbxSelDate2 = Null
    End If
    !tbxSelDate1.Enabled = False: !tbxSelDate2.Enabled = False
  End With
End Sub

Sub YearRangeCheckboxClick()
  With Forms(cstFilterForm)
    !cbxSelYear1.Enabled = !chkYear: !cbxSelYear1 = Null
    !cbxSelYear2.Enabled = !chkYear: !cbxSelYear2 = Null '!chkYear
    If !chkYear Then !cbxSelYear1 = Year(Date)
  End With
End Sub

Sub MonthRangeCheckboxClick()
  With Forms(cstFilterForm)
    !cbxSelMonth1.Enabled = !chkMonth: !cbxSelMonth1 = Null
    !cbxSelMonth2.Enabled = !chkMonth: !cbxSelMonth2 = Null
    If !chkMonth Then
      If !cbxSelQtr1.Enabled Then  'If Qtr enabled, then disable
        !cbxSelQtr1 = Null: !cbxSelQtr2 = Null
        !chkQuarter = 0
        !cbxSelQtr1.Enabled = False
        !cbxSelQtr2.Enabled = False
      End If
    End If
  End With
End Sub

Sub QtrRangeCheckboxClick()
  With Forms(cstFilterForm)
    !cbxSelQtr1.Enabled = !chkQuarter: !cbxSelQtr1 = Null
    !cbxSelQtr2.Enabled = !chkQuarter: !cbxSelQtr2 = Null
    If !chkQuarter Then
      If !cbxSelMonth1.Enabled Then 'If month was enabled, then disable
        !cbxSelMonth1 = Null: !cbxSelMonth2 = Null
        !chkMonth = 0
        !cbxSelMonth1.Enabled = False
        !cbxSelMonth2.Enabled = False
      End If
    End If
  End With
End Sub

Function DateRangeBeforeUpdate(strStartOrEnd As String, strRangeType As String, _
                               strFieldName As String) As Integer
' Accepted values:
'     strStartOrEnd:  start, end.
'     strRangeType:   year, month, quarter, date.
'     strFieldName:   valid date range field on frmFilterform.
' Returns a value for the BeforeUpdate Cnacel variable.
'
  Dim ctl1 As Control, ctl2 As Control
  Dim strOtherField As String, strMsg As String, strTitle As String
  Dim iCancelUpd As Integer
  On Error GoTo DateRangeBeforeUpdateErr

  Select Case strStartOrEnd
  
    Case "start", "end"
      strStartOrEnd = lcase(strStartOrEnd)
      strRangeType = lcase(strRangeType)
      strOtherField = Left(strFieldName, Len(strFieldName) - 1)
      
      With Forms(cstFilterForm)
        Set ctl1 = .Controls(strFieldName)
        If strStartOrEnd = "start" Then
          strOtherField = strOtherField & 2
        Else
          strOtherField = strOtherField & 1
        End If
        Set ctl2 = .Controls(strOtherField)
      End With
        
      If IsNull(ctl2) Then GoTo DateRangeBeforeUpdateExit
      If strStartOrEnd = "start" Then
        If ctl1 > ctl2 Then
          strTitle = "Start " & StrConv(strRangeType, vbProperCase) & " Check"
          strMsg = "Start " & strRangeType & " > End " & strRangeType & vbCrLf & _
                   "To select an open-ended range," & vbCrLf & _
                    "you need to clear the end " & strRangeType & " field."
          iCancelUpd = True
        End If
      Else
        If ctl1 < ctl2 Then
          strTitle = "End " & StrConv(strRangeType, vbProperCase) & " Check"
          strMsg = "End " & strRangeType & " < Start " & strRangeType & vbCrLf & _
                  "To select an up-to-" & strRangeType & " range," & vbCrLf & _
                  "you need to clear the start " & strRangeType & " field."
          iCancelUpd = True
        End If
      End If
      
      If iCancelUpd = True Then
        Beep
        If MsgBox(strMsg & vbCrLf & "Is this what you want to do?", _
                         vbQuestion + vbYesNo, strTitle) = vbYes Then
          ctl2 = Null
          iCancelUpd = False
        End If
      End If

  End Select
  
DateRangeBeforeUpdateExit:
  Set ctl1 = Nothing
  Set ctl2 = Nothing
  DateRangeBeforeUpdate = iCancelUpd
  Exit Function
  
DateRangeBeforeUpdateErr:
  MsgBox "Error: (" & Err & ") " & Err.Description, vbExclamation, cstMDL & ": DateRangeBeforeUpdate"
  Resume DateRangeBeforeUpdateExit
End Function

Sub RefreshActivateMainForm()
  If Nz(Forms(cstFilterForm)!tbxCallingForm, cstMainForm) <> cstMainForm Then
    Call ApplyFormFilter(cstMainForm)
  End If
  Forms(cstMainForm).Requery
  Forms(cstMainForm).SetFocus
End Sub

Sub GetFilteredRecords(strFrmCaller As String)
  On Error GoTo GetFilteredRecordsErr
  
  Call ApplyFormFilter(strFrmCaller)
  Forms(strFrmCaller).SetFocus
  
GetFilteredRecordsExit:
  Exit Sub
  
GetFilteredRecordsErr:
  MsgBox "Error: (" & Err & ") " & Err.Description, vbExclamation, cstMDL & " GetFilteredRecords"
  Resume GetFilteredRecordsExit
End Sub

Sub ApplyFormFilter(strFrmCaller As String)
' To requery the underlying form qry or (cbx data source on the Deal Detail form),
' and toggle the label/button colors to show whether a filter is applied.
  On Error GoTo ApplyFormFilterErr
  If Not IsLoaded(strFrmCaller) Then Exit Sub
  
  Call ResetFilterClues(strFrmCaller)
  If strFrmCaller = cstDealForm Then
    If Not IsNull(Forms(strFrmCaller)!cbxSelDealNum) Then
      Forms(strFrmCaller)!cbxSelDealNum.Requery
    End If
  End If
  Forms(strFrmCaller).Requery
  Exit Sub
  
ApplyFormFilterErr:
  MsgBox "Error: (" & Err & ") " & Err.Description, vbExclamation, cstMDL & " ApplyFormFilter"
End Sub

Sub ResetFilterClues(strFrm As String)
  Dim strCountCaption As String
  On Error GoTo ResetFilterCluesErr
  
  If Not IsLoaded(cstFilterForm) Then DoCmd.OpenForm cstFilterForm, , , , , acHidden
  Forms(cstFilterForm)!txtEditOrder = ""

  With Forms(strFrm)
    If FormFilterIsSet Then
      !lblFilter.BackColor = cstAAColorLight
      !lblFilter.ForeColor = vbWhite
      strCountCaption = "Filtered deal(s)"
    Else
      !lblFilter.BackColor = vbWhite
      !lblFilter.ForeColor = cstAAColorLight
      strCountCaption = "Unfiltered deals"
    End If
  End With
  If strFrm = cstMainForm Then
    Forms(cstMainForm).Requery
    Forms(cstMainForm)!lblCount.Caption = strCountCaption
  End If
  Exit Sub
  
ResetFilterCluesErr:
  MsgBox "Err: " & Err.Number & " - " & Err.Description, vbExclamation, _
         cstMDL & " ResetFilterClues"
End Sub

Function DateCtlAfterUpdate(iCtlNum As Integer)
  Dim ctl1 As Control, ctl2 As Control
  Dim chk As CheckBox
  On Error GoTo DateCtlAfterUpdateErr
  
  With Forms(cstFilterForm)
    Select Case iCtlNum
      Case 1
        Set ctl1 = !cbxSelYear1
        Set ctl2 = !cbxSelYear2
        Set chk = !chkYear
        
      Case 2
        Set ctl1 = !cbxSelMonth1
        Set ctl2 = !cbxSelMonth2
        Set chk = !chkMonth
      
      Case 3
        Set ctl1 = !cbxSelQtr1
        Set ctl2 = !cbxSelQtr2
        Set chk = !chkQuarter
        
      Case 4
        Set ctl1 = !tbxSelDate1
        Set ctl2 = !tbxSelDate2
        Set chk = !chkRange
    End Select
  End With
  
  Call AddToEditList(iCtlNum)
  chk = (Not IsNull(ctl1) Or Not IsNull(ctl2))

DateCtlAfterUpdateExit:
  Set ctl1 = Nothing
  Set ctl2 = Nothing
  Set chk = Nothing
  Exit Function
  
DateCtlAfterUpdateErr:
  MsgBox "Err: " & Err.Number & " - " & Err.Description, vbExclamation, cstMDL & " DateCtlAfterUpdate"
  Resume DateCtlAfterUpdateExit
End Function

Function AddToEditList(idxCtl As Integer)
  Dim frm As Form
  
  Set frm = Forms(cstFilterForm)
  With frm
    If Len(!txtEditOrder) = 0 Then
      !txtEditOrder = ";" & idxCtl
    Else
      If CInt(Right$(!txtEditOrder, 1)) <> idxCtl Then
        !txtEditOrder = !txtEditOrder & ";" & idxCtl
      End If
    End If
  End With
  Set frm = Nothing
  
End Function

Function UndoLastEdits()
  Dim frm As Form
  Dim ctl As Control
  Dim strOrder As String
  Dim idx As Integer
  Dim lngDelim As Long
  
  Set frm = Forms(cstFilterForm)
  If Len(frm!txtEditOrder) = 0 Then GoTo UndoLastEditsExit
  
  Do While Len(frm!txtEditOrder) > 1
    strOrder = ""
    strOrder = StrReverse(frm!txtEditOrder) ' last edit is at begining of string
    lngDelim = InStr(1, strOrder, ";")      ' find first delimiter
    idx = CInt(StrReverse(Left$(strOrder, lngDelim - 1)))
    If Len(strOrder) > 1 Then
      strOrder = StrReverse(Mid$(strOrder, lngDelim + 1))
      frm!txtEditOrder = strOrder 'update ctl
    Else
      frm!txtEditOrder = ""
    End If
    
    For Each ctl In frm.Controls
      If ctl.ControlType = acComboBox Or ctl.ControlType = acTextBox Then
        If ctl.TabIndex = idx Then
          Call ClearCtlWithRightMouseDown(ctl, acRightButton)
          Exit For
        End If
      End If
    Next ctl
  Loop
  
UndoLastEditsExit:
  frm.Visible = False
  Set ctl = Nothing
  Set frm = Nothing
  
End Function

Public Function GetFilterName() As String
 GetFilterName = IIf(IsLoaded("frmFilterForm"), _
                     IIf(Len([Forms]![frmFilterForm]![tbxFilter]) = 0, " No filter", _
                          [Forms]![frmFilterForm]![tbxFilter]), " [FilterForm not used]")
End Function
