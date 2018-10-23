Attribute VB_Name = "MDL_SeriesForm"
Option Compare Database
Option Explicit
'================================================================================
'
' MDL SeriesForm  Dec-04-03
'   Change Dec-04-03 SaveNewSeries:cannot set frm to Visible=False when it has focus
'   Change Nov-07-03 12:30: Removed control name as param from ResetAmtsGrid function
'================================================================================
Const cstMDL = "SeriesForm"
'
Const cstQryUpdate = "qrySeriesUpdate"
Public dteLastColDate As Date ' Usage:
' Series period update: if upd date < dteLastColDate, stop col entry, else ask if new col.
Public Const cstRowHeight = 255   '(260 * iRows +2)/1440 = Series Grid height in inches
Public Const cstDateDisp = "mm/dd/yyyy"
Public Const cstCurrDisp = "#,##0;(#,##0)"

Sub RefreshMainGrids()
  Dim dbDAO As DAO.Database
  Dim rst As DAO.Recordset

  Dim dte As Date
  Dim frmMain As Form, sfrm As Form
  Dim lngGridH As Long, lngRecs As Long, lngOutcome As Long
 
  Const cstGridDataMTBLQry = "qryGridSeriesDataMTBL"
  Const cstTblGridData = "tblGridSeriesData"
  Const cstQryGrid = "qryGridSeriesData"
  Const cstGridCtlHeight = 5.12 * 1440
  
  lngR = 0: lngC = 0: lngGridH = 0: lngRecs = 0
  
  On Error GoTo RefreshGridsErr
  Screen.MousePointer = 11 'busy
  
  If IsNull(Forms(cstFRM_Main)!cbxSelForecast) Then Exit Sub

  Set frmMain = Forms(cstFRM_Main)
  With frmMain
    !lblProcessing.Left = !lblBtn2.Left
    !lblProcessing.Visible = True
    lngCurrentComp = !cbxSelComp
    lngCurrentForecast = !cbxSelForecast
  End With

RecreateTable:    'cstTblGridData
  Set dbDAO = CurrentDb
  
  lngRecs = ExecQuery(dbDAO, cstGridDataMTBLQry, lngCurrentForecast)
  If lngRecs = 0 Then 'add first col
    dte = Nz(DLookup("[dteForecastDate]", "tblForecasts", "[lngForecastID]=" & _
                                                           lngCurrentForecast), Date)
    lngOutcome = AddSeries(lngCurrentForecast, dte)
    If lngOutcome = 0 Then   ' Add default series using frc date
      GoTo RecreateTable
    Else
      Err.Raise (lngOutcome)
    End If
  End If
  
  ' Resize tbl according to company-defined fields if necessary: (including all required fields)
  Call ProcessCompFields(dbDAO, cstTblGridData)
  
  ' Open rst against that tbl:
  Set rst = dbDAO.OpenRecordset(cstQryGrid)
  
  ' Populate series grid:
  Set sfrm = frmMain!sfrmAny.Form    'ref subform
  Call FillGridT(sfrm!ocxFlxGridSeries, rst, lngR, lngC)
  
  ' Cleanup
  rst.Close
  dbDAO.Close
  
  Call ResetAmtsGrid(lngR, sfrm)

  lngGridH = ((cstRowHeight + (2 * lngR)) * lngR) + (lngR * 4)
  With sfrm
    If lngGridH < cstGridCtlHeight Then  'resize
      !ocxFlxGridSeries.Height = lngGridH
      !ocxFlxGridEdit.Height = lngGridH
    End If
      
    With !ocxFlxGridSeries
      If lngC > 0 Then .LeftCol = lngC  ' Scroll to last col
      .TopRow = 1
      .RowSel = .Row
      .ColSel = .col 'remove prev selection
      dteLastColDate = CDate(.TextMatrix(0, lngC))
    End With
    !ocxFlxGridEdit.TopRow = 0
  End With

RefreshGridsExit:
  frmMain!lblProcessing.Visible = False
  Screen.MousePointer = 0
  Set sfrm = Nothing
  Set frmMain = Nothing
  Set rst = Nothing
  Set dbDAO = Nothing
  Exit Sub
  
RefreshGridsErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , cstMDL & ":RefreshMainGrids"
  Resume RefreshGridsExit
End Sub

Public Sub FillGridT(ocxGrid As Control, rstInputData As DAO.Recordset, _
                                         lngRows As Long, lngCols As Long)
' Called by RefreshMainGrids.
' Performs transposition of Row & Col.
' The number of fields used must correspond to the number of columns in the ocxGrid control.
'
  Dim staNames() As String  ' Array for fldData display (column initially, then rows) names.
  Dim lngRecs As Long, lngFlds As Long, lngIdxFld As Long, lngIdxRow As Long
  Dim lngColW As Long, lngDataColW As Long, lngCol As Long
  Dim i As Integer, iFields As Integer
  Dim varVal, varData()       ' For GetRows method
  Dim varTrspData             ' For Transposed GetRows array
  Const cstSeriesColW = 1125
  '------------------------------------------------------------------------------------
  lngIdxFld = 0: lngIdxRow = 0: lngRecs = 0: lngFlds = 0: i = 0: lngColW = 0: lngCol = 0
  '------------------------------------------------------------------------------------
  ' Reset to ctl default row & col numbers:
  ocxGrid.Rows = 2: ocxGrid.Cols = 2: lngRows = 1: lngCols = 1
  '------------------------------------------------------------------------------------
  On Error GoTo FillGridTErr
  
  If rstInputData.AbsolutePosition = -1 Then GoTo FillGridTExit
  rstInputData.MoveLast
  rstInputData.MoveFirst
  
  For i = 0 To rstInputData.Fields.Count - 1
    ReDim Preserve staNames(i)
    staNames(i) = rstInputData.Fields(i).Name  'Properties("Caption")
  Next i
   
  lngRecs = rstInputData.RecordCount         ' Needed for GetRows
  varData = rstInputData.GetRows(lngRecs)    ' Returns a 2-dimensional array(Field, Record)
  varTrspData = Transpose2DArray(varData) ' Returns a 2-dimensional array(row, fld)
  lngRecs = UBound(varTrspData, 1)
  lngFlds = UBound(varTrspData, 2)
  
  ' Create number of columns to equal, at least, number of fields.
  If ocxGrid.Cols < lngRecs + 2 Then
     ocxGrid.Cols = lngRecs + 2
     '1+1=1 for idx offset + 1 because 1st colum will holds "row name"
  End If
  
  i = 0
  With ocxGrid
    For lngIdxFld = 0 To lngFlds
      If lngIdxFld > 1 Then .AddItem ""
    
      For lngIdxRow = 0 To lngRecs + 1
        
        If lngIdxRow = 0 Then 'column 0
          .TextMatrix(lngIdxFld, 0) = staNames(lngIdxFld) '
          ' Get col width to fit largest field name in the list:
          lngCol = (Len(staNames(lngIdxFld)) * 90)
          If lngColW < lngCol Then lngColW = lngCol
        Else
        
          varVal = 0
          i = lngIdxRow - 1
          varVal = varTrspData(i, lngIdxFld)
          
          If Not IsNull(varVal) Then
            If IsNumeric(varVal) Then
              varVal = Nz(varVal, 0)
              .TextMatrix(lngIdxFld, lngIdxRow) = Format(varVal, cstCurrDisp)
            Else
              .TextMatrix(lngIdxFld, lngIdxRow) = varVal
            End If
          Else
            .TextMatrix(lngIdxFld, lngIdxRow) = Format(0, cstCurrDisp)
          End If
          ' Remove format on Months row:
          If lngIdxFld = 1 Then .TextMatrix(lngIdxFld, lngIdxRow) = varVal
        End If
        '
        .ColAlignment(lngIdxRow) = flexAlignRightCenter  'format col
        .ColWidth(lngIdxRow) = cstSeriesColW   'resize data cols
      
      Next lngIdxRow
      
      .RowHeight(lngIdxFld) = cstRowHeight
    Next lngIdxFld
    .ColWidth(0) = lngColW      'cstSeriesColW '               ' Resize firstData column
    .ColAlignment(0) = flexAlignLeftCenter     ' Set firstData column's alignment
  End With
    
  'Assign final row/col num to output vars:
  lngRows = lngFlds + 1: lngCols = lngRecs + 1
  
FillGridTExit:
  Erase varData
  Erase staNames
  Exit Sub

FillGridTErr:
  If Not (rstInputData Is Nothing) Then rstInputData.Close
  Set rstInputData = Nothing
  lngRows = 0: lngCols = 0
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , cstMDL & ":FillGridT"
  Resume FillGridTExit
End Sub

Function ResetAmtsGrid(lngRows As Long, frm As Form)
  Dim r As Integer
  Dim ocxAmts As Control
  'Const cstEditingGrid = "ocxFlxGridEdit"
  
  Set ocxAmts = frm!ocxFlxGridEdit '.Controls( cstEditingGrid )
  ' Refresh Amts grid:
  ocxAmts.Rows = 1: ocxAmts.Cols = 1
  
  If lngRows = 0 Then Exit Function
  With ocxAmts
    For r = 0 To lngRows - 1
      If r = 0 Then
        .TextMatrix(0, 0) = "<end date>"  ' Populate grid header
      Else
        If r > 0 Then .AddItem ""
        If r = 1 Then 'cell for month num
          .TextMatrix(r, 0) = 0
        Else
          .TextMatrix(r, 0) = Format(0, cstCurrDisp)
        End If
      End If
      .RowHeight(r) = cstRowHeight
    Next r
    .ColWidth(0) = 1290
    .ColAlignment(0) = flexAlignRightCenter
  End With
  Set ocxAmts = Nothing
  
End Function

Public Function UpdateSeries(dteSeriesDate As Date, iMonth As Integer, _
                                varNewData As Variant) As Integer
' Purpose:  To update the series displayed in the data grid with the amendments in the
'           edit grid.
' Arg:      varNewData is a 3-dim array:
'             column(0)=fld internal name;
'             column(1)=new values to update all the editable fields in the Series table;
'             column(2)=field display name to avoid another lookup if msgbox is used.
' Usage:    Called by user
' Output:   Returned value = err.num or 0 if no error
'
  Dim dbDAO As DAO.Database
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim strFld$
  Dim varLookup As Variant
  Dim lngDataFlds As Long, lngVal As Long
  Dim d As Integer, iErr As Integer, t As Integer, f As Integer
  iErr = 0
  
  On Error GoTo UpdateSeriesErr
  If IsNull(varNewData) Then
    iErr = 99  ' Debug.Print "IsNull(varNewData)", iErr
    GoTo UpdateSeriesErr
  End If
  
  lngDataFlds = UBound(varNewData, 1) '(all dims have same size)
  If lngDataFlds = 0 Then
    iErr = 99  'Debug.Print "lngDataFlds = 0 ", iErr
    GoTo UpdateSeriesErr
  End If
  
  Set dbDAO = CurrentDb
  Set qdf = dbDAO.QueryDefs(cstQryUpdate)
  qdf.Parameters(0) = lngCurrentForecast
  qdf.Parameters(1) = dteSeriesDate
  qdf.Parameters(2) = iMonth
  Set rst = qdf.OpenRecordset(dbOpenDynaset)
  If rst.AbsolutePosition = -1 Then
    iErr = 999
    GoTo UpdateSeriesErr
  End If
    
  rst.MoveLast
  rst.Edit
  For d = 0 To lngDataFlds        '= UBound(varNewData, 1)
    strFld$ = varNewData(0, d)   'save fld internal name e.g. curCash
    lngVal = varNewData(1, d)
    If d > 1 Then 'currency fields
      If lngVal >= 1000000000 Then  'over 1 billion
        varLookup = DLookup("[txtDispName]", "tlkpAllFields", "[txtFldTblName]='" & _
                            strFld$ & "'")
        If MsgBox("Amounts are stored in thousands of dollars." & vbCrLf & _
                  "You have entered $" & lngVal & " for " & varLookup & "." & vbCrLf & _
                  "Is this correct?", vbYesNo, "Billion Amount Check") = vbNo Then
          lngVal = lngVal / 1000
          lngVal = CCur(InputBox("Either accept the proposed correction, " & _
                                 "or edit it and click OK.", _
                                 "Billion Amount Correction", lngVal))
        End If
      End If
    End If
    rst.Fields(strFld$) = lngVal 'update fld in rst, e.g. rst.fld("curCash")
  Next d
  rst.Update
  
  ' Clean up
  rst.Close
  qdf.Close
  dbDAO.Close
  
UpdateSediesExit:
  UpdateSeries = iErr
  Set rst = Nothing
  Set qdf = Nothing
  Set dbDAO = Nothing
  Exit Function
  
UpdateSeriesErr:
  If iErr <> 0 Then
    If iErr = 99 Then
      Debug.Print "UpdateSeries error: the array of updated values is empty."
    ElseIf iErr = 999 Then
      Debug.Print "UpdateSeries error: empty recordset."
    End If
    GoTo UpdateSediesExit
  ElseIf Err <> 0 Then
    iErr = Err.Number
    MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstMDL & ":UpdateSeries"
    Resume UpdateSediesExit
  End If
End Function

Function AddSeries(lngForcID As Long, Optional dteForecastDate As Variant) As Long
  Dim dte As Date
  Dim dbDAO As DAO.Database
  Dim rst As DAO.Recordset
  AddSeries = 0
  On Error GoTo AddSeriesErr
  
  dte = GetMonthEndDate(dteForecastDate)

  Set dbDAO = CurrentDb
  Set rst = dbDAO.OpenRecordset("Select * From tblSeriesData")
  With rst
    .MoveLast
    .AddNew
    !lngForecastID = lngForcID
    !dtePeriodEndDate = dte
    !intMonths = 1
    .Update
  End With
  rst.Close
  dbDAO.Close
  
AddSeriesExit:
  Set rst = Nothing
  Set dbDAO = Nothing
  Exit Function
  
AddSeriesErr:
  AddSeries = Err.Number
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstMDL & ":AddSeries"
  Resume AddSeriesExit
End Function

Public Function DeleteSeries(ByVal lngForecastID As Long, dteSeriesDate As Date, _
                                            iMon As Integer) As Integer
' Purpose:  To delete the series displayed in the data grid
' Usage:    Called by user
' Output:   Returned value = err.num or 0 if no error
  Dim dbDAO As DAO.Database
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim strSQL As String
  Dim iErr As Integer
  
  iErr = 0
  On Error GoTo DeleteSeriesErr

  Set dbDAO = CurrentDb
  Set qdf = dbDAO.QueryDefs(cstQryUpdate)
  qdf.Parameters(0) = lngForecastID
  qdf.Parameters(1) = dteSeriesDate
  qdf.Parameters(2) = iMon
 
  Set rst = qdf.OpenRecordset(dbOpenDynaset)
  If rst.AbsolutePosition <> -1 Then
    rst.MoveLast
    rst.Delete
  End If
  'Cleanup
  rst.Close
  qdf.Close
  dbDAO.Close
  
DeleteSeriesExit:
  DeleteSeries = iErr
  Set rst = Nothing
  Set qdf = Nothing
  Set dbDAO = Nothing
  Exit Function
  
DeleteSeriesErr:
  iErr = Err.Number
  Debug.Print "DeleteSeries Err (" & Err.Number & "): " & Err.Description
  Resume DeleteSeriesExit
End Function

Public Function SaveNewSeries()
  Dim frm As Form
  On Error GoTo SaveNewSeriesErr

  Set frm = Forms(cstFRM_NewSer)
  With frm
    If .Dirty Then
      frm!intMonths = frm!opgPeriod.Value
      DoCmd.RunCommand acCmdSaveRecord
    End If
    .Modal = False
    Forms(cstFRM_Main)!sfrmAny.SourceObject = cstSFRM_Series
  End With
  Set frm = Nothing
  DoCmd.Close acForm, cstFRM_NewSer
    
SaveNewSeriesExit:
  Set frm = Nothing
  Exit Function
    
SaveNewSeriesErr:
  MsgBox "SaveNewSeries Err (" & Err.Number & "): " & Err.Description
  Resume SaveNewSeriesExit
End Function

