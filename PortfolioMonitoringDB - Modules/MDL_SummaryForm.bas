Attribute VB_Name = "MDL_SummaryForm"
Option Compare Database
Option Explicit
'================================================================================
'
' MDL SummaryForm Sep-16-03 17:55
'
'================================================================================
Const cstMDL = "SummaryForm"
'
Dim lngFrcType As Long
Dim strTypeDesc As String

Dim lngRevM1 As Long, lngRevM2 As Long
Dim lngGProfM1 As Long, lngGProfM2 As Long
Dim lngEbitdaM1 As Long, lngEbitdaM2 As Long
Dim lngAdjEbitdaM1 As Long, lngAdjEbitdaM2 As Long
Dim lngIntExpM1, lngIntExpM2 As Long

Dim lngEbitdaLTM1 As Long, lngAdjEbitdaLTM1 As Long
Dim lngEbitdaLTM2 As Long, lngAdjEbitdaLTM2 As Long
Dim lngEbitdaLTM3 As Long, lngAdjEbitdaLTM3 As Long
Dim lngIntExpLTM1 As Long, lngIntExpLTM2 As Long, lngIntExpLTM3 As Long

Dim lngSrSubM1 As Long, lngSrSubM2 As Long, lngSrSubM3 As Long
Dim lngTotDebtM1 As Long, lngTotDebtM2 As Long, lngTotDebtM3 As Long
Dim lngTotCapM1 As Long, lngTotCapM2 As Long, lngTotCapM3 As Long

Dim lngLastCapRow As Long, lngFirstCapRow As Long  'to check when filling WC section
Dim lngLastWCRow As Long, lngTotDebtRow As Long, lngTotEqRow As Long

Public Const cstFinalTbl = "tblSummaryFinal"  'transposed table (CreateTransposedSFTable)
' col width- Summary grid
Public Const cstNamesColW = 1700  ': initial
Public Const cstLargeColW = 1400
Public Const cstPeriodDescColW = 1060
Public Const cstSeriesColW = 945
Public Const cstChangeColW = 895
Public Const cstTwips = 150 ' 140= when not bold
Const cstChangeDisp = "#.#0%"
Const cstPctDisp = "#.#%"
Const cstOffsetCol = 4
Const cstCent = 100

Sub RefreshSummaryGrid()
  Dim dbDAO As DAO.Database
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset, rstAdd As DAO.Recordset
  Dim lngRecAffected As Long
  Dim frmMain As Form, frm As Form
  Dim strAPNDQry As String, strSummaryTotalSQL As String, strCompany As String
  Dim strSection As String ', strCol As String
  Dim dteSelected As Date, dtePrev As Date, dteMulti As Date
  Dim varQryParams() As Variant
  Dim blnWarn As Boolean
  Dim lngTotApnd As Long, lngActualApnd As Long, lngBudgetApnd As Long
  Dim strSQL As String
  
  Dim f As Integer, r As Integer, m As Integer, n As Integer, iMonths As Integer
  Dim sglMulti As Single
   
  ' Queries, in order of use:
  Const cstQrySeriesDataMTBL = "qrySF-SeriesIniMTBL"        '1
  Const cstQrySeriesDataMulti = "qrySF-Series-MultiMonths"  '2, select qry
  Const cstSummaryGridMTBLQry = "qrySF-SummaryIni-MTBL"     '3
  Const cstQryAPNDPrefix = "qrySF-APND-"                    '4, called 9 times
  Const cstQryGrid = "qrySF-SummaryData"
  ' Tables:
  Const cstTblSeriesData = "tblSF-SeriesData"     'created by "qrySF-SeriesIniMTBL"
  Const cstTblSummaryData = "tblSummaryFormData"  'created by "qrySF-SummaryIni-MTBL"
  ' Warning:
  Const cstNoActual = "PROBLEM: A Budget forecast exists while there is no Actual forecast."
  '-----------------------------------------------------------------------
  'dteSelected = 0: dtePrev = 0: lngTotApnd = 0
  'r = 0: m = 0: n = 0: iMonths = 0: sglMulti = 0:
  'blnWarn = False
  'strCompany = ""
  '-----------------------------------------------------------------------
  'On Error GoTo RefreshSummaryGridErr
  
  Screen.MousePointer = 11 'busy
  
  Set dbDAO = CurrentDb
  
  Set frmMain = Forms(cstFRM_Main)
  Set frm = frmMain!sfrmAny.Form    'ref subform
  frm!lblWarn.Visible = True
  frm!lblWarn.Caption = ""
  
  If Not CoHasDefaultBudget(frmMain!cbxSelComp) Then
    frm!lblWarn.Caption = "PROBLEM: No default forecast, no Summary."
    GoTo RefreshSummaryGridExit
  End If
  
  frmMain!lblProcessing.Left = frmMain!lblBtn3.Left
  frmMain!lblProcessing.Visible = True
  dteSelected = frm!cbxSelPeriodEnd
  sglMulti = frm!MultiUpd
  strCompany = frm!txtCurrentCompany
  lngFrcType = frmMain!cbxSelForecast.Column(2)
  frm!ocxFlxGridForm.MousePointer = flexHourglass
  DoEvents
  
  strTypeDesc = DLookup("[txtForecastType]", "tlkpForecastType", _
                        "[lngForecastTypeID]=" & lngFrcType)

  ' Recreate initial, stand-alone series data tbl; Run the initial mtbl qry
  ReDim varQryParams(1)                           'OK
  varQryParams(0) = lngFrcType
  varQryParams(1) = lngCurrentComp
  lngRecAffected = ExecQuery(dbDAO, cstQrySeriesDataMTBL, varQryParams(0), varQryParams(1))
  Erase varQryParams
  If lngRecAffected = 0 Then
    MsgBox cstNoFrcForThisComp, vbExclamation, "No data"
    GoTo RefreshSummaryGridExit
  End If
  
  ' Get the records where intMonths>1
  Set rst = dbDAO.QueryDefs(cstQrySeriesDataMulti).OpenRecordset
  If rst.AbsolutePosition <> -1 Then
    rst.MoveLast
    rst.MoveFirst
    strSQL = "SELECT [tblSF-SeriesData].* FROM [tblSF-SeriesData];"
    Set rstAdd = dbDAO.OpenRecordset(strSQL, dbOpenDynaset)
    ' Convert each one to monthly using new end dates
    For r = 1 To rst.RecordCount
      rstAdd.MoveLast
      iMonths = rst!intMonths
      dteMulti = rst!dtePeriodEndDate
      Debug.Print "RefreshSummaryGrid iMonths: " & iMonths & ";  dteMulti: " & dteMulti
      
      For m = 1 To iMonths
        ' Calc new date
        dtePrev = DateSerial(Year(dteMulti), Month(dteMulti) + m + 1 - iMonths, 0)

        rstAdd.AddNew 'copy all values
        For f = 0 To rstAdd.Fields.Count - 1
          If f < 5 Then 'first 6 fields are the non currency fields:
            'lngForecastID, lngForecastTypeID, intDefaultBudget, dtePeriodEndDate, intMonths
            rstAdd(f) = rst(f)
          Else
            rstAdd(f) = rst(f) / iMonths
          End If
        Next f
        ' Reset the fields that have new data:
        rstAdd!dtePeriodEndDate = dtePrev
        rstAdd!intMonths = 1
        rstAdd.Update
      Next m
      rst.MoveNext
    Next r
    
    dbDAO.TableDefs.Refresh
  End If
  rst.Close
  Set rst = Nothing
  
  Erase varQryParams
  
  ' Recreate independent summary form selection tbl: (no parmas)
  Call ExecQuery(dbDAO, cstSummaryGridMTBLQry, varQryParams)  'varQryParams not set

  ReDim varQryParams(0)
  varQryParams(0) = dteSelected
  ' Populate with APND qries for each kind of forecast, using the selected date
  For m = 1 To 3   'number of form sections: Month End - YTD - LTM
    lngTotApnd = 0: lngActualApnd = 0: lngBudgetApnd = 0: strSection = "" 'reset
    For n = 1 To 3 'number of forecasts/date calculations: Last yr, Current, Budget
      strAPNDQry = "":
      strAPNDQry = cstQryAPNDPrefix & m & "-" & n
      lngTotApnd = ExecQuery(dbDAO, strAPNDQry, varQryParams(0))  ', varQryParams(1))
      Select Case n
        Case 2
          lngActualApnd = lngTotApnd
        Case 3
          lngBudgetApnd = lngTotApnd
      End Select
    Next n
    
    If lngFrcType = 1 Then
      If lngActualApnd = 0 And lngBudgetApnd <> 0 Then
        Select Case m
          Case 1
            strSection = " Section: End Of Month (EOM)."
          Case 2
            strSection = " Section: Year To Date (YTD)."
          Case 3
            strSection = " Section: Last Twelve Months (LTM)."
        End Select
        blnWarn = True
      End If
      If blnWarn Then Exit For
    End If
  Next m
  dbDAO.TableDefs.Refresh
  
  With frm!ocxFlxGridForm
    .Rows = 1
    .Clear
    .AllowUserResizing = flexResizeColumns
    .FillStyle = flexFillSingle
    .FontBold = False
    .ColWidth(0) = cstNamesColW
    .CellAlignment = flexAlignLeftCenter
  End With
      
  If blnWarn Then
    beep
    frm!lblWarn.Visible = True
    frm!lblWarn.Caption = cstNoActual & strSection
    GoTo RefreshSummaryGridExit
  End If
  
  ReDim varQryParams(0)
  'Append 3 future columns for processing:
  Call ExecQuery(dbDAO, "qrySF-APND-AcctgOrder", varQryParams(0)) 'varQryParams not set (there are none)
  Call ExecQuery(dbDAO, "qrySF-APND-FldOrder", varQryParams(0))
  Call ExecQuery(dbDAO, "qrySF-APND-Priority", varQryParams(0))
    
  ' Resize tbl according to company-defined fields if necessary:
  Call ProcessCompFields(dbDAO, cstTblSummaryData)
    
  ' Form SQL to obtain a GroupBy qry that will use only the comp defined fields.
  strSummaryTotalSQL = GetSummaryDataSQL(lngCurrentComp)
  'Debug.Print "strSummaryTotalSQL: " & vbCrLf & strSummaryTotalSQL
  
  Set rst = dbDAO.OpenRecordset(strSummaryTotalSQL)
  Call CreateTransposedSFTable(rst, cstFinalTbl)
  rst.Close
  Set rst = Nothing
  
  ' Populate series grid:
  Call FillSummaryGrid(frm!ocxFlxGridForm, dteSelected, strCompany, sglMulti)
      
RefreshSummaryGridExit:
  Screen.MousePointer = 0
  Erase varQryParams
  frmMain!lblProcessing.Visible = False
  frm!ocxFlxGridForm.MousePointer = flexDefault
  Set frmMain = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbDAO = Nothing
  Exit Sub
  
RefreshSummaryGridErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , cstMDL & ":RefreshSummaryGrid"
  Resume RefreshSummaryGridExit
End Sub

Sub FillAddBlankRow(ocxGrid As Control, lngColStart As Long, lngColEnd As Long)
  Dim lngRo As Long, lngCo As Long
  
  With ocxGrid
    .AddItem ""
    lngRo = .Rows - 1
    .TextMatrix(lngRo, lngColStart) = " "
  End With
End Sub

Function CreateTransposedSFTable(rstSource As DAO.Recordset, strTarget As String)
 ' Creats a transposed tbl to be used to fill the Summary Form grid.
  Dim tdfNew As DAO.TableDef
  Dim fldNew As DAO.Field
  Dim idxNew As DAO.Index
  Dim qdf As DAO.QueryDef
  Dim rstCompFlds As DAO.Recordset
  Dim rstTarget As DAO.Recordset
  Dim strFind As String
  Dim i As Integer, j As Integer
  Dim lngVal As Long
  i = 0: j = 0: strFind = "": lngVal = 0
  
  On Error GoTo CreateTransposedTableErr
  'Debug.Print "Start CreateTransposedSFTable"
  rstSource.MoveLast
  rstSource.MoveFirst
  
  ' Create a new table to hold the transposed data.
  ' Create a field for each record in the original table.
  If IsTableInDB(strTarget) Then CurrentDb.TableDefs.Delete strTarget
  Set tdfNew = CurrentDb.CreateTableDef(strTarget)

  ' First fld is text
  Set fldNew = tdfNew.CreateField(CStr(0), dbText) 'FormCat
  tdfNew.Fields.Append fldNew
  ' All others are long
  For i = 0 To rstSource.RecordCount - 1
    Set fldNew = tdfNew.CreateField(CStr(i + 1), dbLong) 'or dbSingle?
    tdfNew.Fields.Append fldNew
    rstSource.MoveNext
  Next i
  Set fldNew = Nothing
  
  Set idxNew = tdfNew.CreateIndex("Key")
  With idxNew
     .Name = "Key"
     .Primary = True
     .Unique = False
     .Required = True
     .IgnoreNulls = False
  End With
 
  ' Create 3 index fields with the same name as a table field (1,2,3),then append it to the index.
  '1=AcctgOrder; 2=FldOrder, 3=Priority
  For i = 1 To 3
    Set fldNew = idxNew.CreateField(CStr(i))
    idxNew.Fields.Append fldNew
    Set fldNew = Nothing
  Next i
  
  ' Append the index to the TableDef.
  tdfNew.Indexes.Append idxNew
  Set idxNew = Nothing
  ' Append the table to the TableDefs.
  CurrentDb.TableDefs.Append tdfNew
  Set tdfNew = Nothing
  
  ' Open cstQryCoFlds to lookup the ordering:
  Set qdf = CurrentDb.QueryDefs(cstQryCoFlds)
  qdf.Parameters(0) = lngCurrentComp
  Set rstCompFlds = qdf.OpenRecordset
  rstCompFlds.MoveLast
  rstCompFlds.MoveFirst
  
  rstSource.MoveFirst
  ' Open the new table and fill the first field with field names from the original table.
  Set rstTarget = CurrentDb.OpenRecordset(strTarget)
  ' Fill each column of the new table with a record from the original table.
  For j = 1 To rstSource.Fields.Count - 1
    rstTarget.AddNew
    
    For i = 0 To rstTarget.Fields.Count - 1
    
      With rstTarget
        If i = 0 Then
          .Fields(i) = rstSource.Fields(j).Name
        Else
          If j > 0 Then
            If Len(rstSource.Fields(j) & "") = 0 Then
              lngVal = 0
            Else
              lngVal = CLng(rstSource.Fields(j))
            End If
            .Fields(i) = lngVal
            
            If j > 2 Then
              strFind = "[txtDispName]='" & rstSource.Fields(j).Name & "'"
              If i >= 1 And i <= 3 Then
                rstCompFlds.FindFirst strFind
                .Fields(1) = rstCompFlds!AcctgOrder
                .Fields(2) = rstCompFlds!FldOrder
                .Fields(3) = rstCompFlds!Priority
                strFind = ""
              End If
            End If
          End If
          rstSource.MoveNext
        End If
      End With
      
    Next i
    rstTarget.Update
    rstSource.MoveFirst
  Next j
  rstCompFlds.Close
  qdf.Close
  rstTarget.Close
  
  CurrentDb.TableDefs.Refresh
  'Debug.Print "End CreateTransposedSFTable"
  
CreateTransposedTableExit:
  Set fldNew = Nothing
  Set idxNew = Nothing
  Set tdfNew = Nothing
  Set qdf = Nothing
  Set rstTarget = Nothing
  Set rstCompFlds = Nothing
  Exit Function
  
CreateTransposedTableErr:
  Select Case Err
    Case 3010
       MsgBox "The table " & strTarget & " already exists."
    Case 3078
       MsgBox "The rst doesn't exist."
    Case Else
       MsgBox Err & " " & Err.Description, vbExclamation, cstMDL & ": CreateTransposedTable"
  End Select
  Resume CreateTransposedTableExit

End Function

Sub FillHeaderTop(dbDAO As DAO.Database, ocxFlexGrid As Control, strDate As String)
' Fills first header & sizes cols.
  Dim c As Integer, r As Integer, m As Integer
  
  'Initialize cumulative var:
  lngSrSubM1 = 0: lngSrSubM2 = 0: lngSrSubM3 = 0
  lngTotDebtM1 = 0: lngTotDebtM2 = 0:  lngTotDebtM3 = 0:
  lngTotCapM1 = 0:  lngTotCapM2 = 0:  lngTotCapM3 = 0:

  With ocxFlexGrid '(R, C)
    .FillStyle = flexFillSingle 'format each cell separately

    ' Fill hdr row 0
    For c = 1 To 14
      .FillStyle = flexFillSingle
      .Row = 0
      .col = c
  
      .CellAlignment = flexAlignCenterCenter
      .CellFontBold = True
        
      Select Case c
        Case 1, 3, 8, 13
          .TextMatrix(0, c) = ""
          .ColWidth(c) = cstSeriesColW
        Case 2
          .TextMatrix(0, c) = "EOM " & strDate
          .ColWidth(c) = cstPeriodDescColW
        Case 4, 9, 14
          .TextMatrix(0, c) = "Change %"
          .ColWidth(c) = cstChangeColW
        Case 5, 15  'change cols
          .TextMatrix(0, c) = ""
          .ColWidth(c) = cstChangeColW
        Case 6, 10
          .TextMatrix(0, c) = ""
          .ColWidth(c) = cstLargeColW
        Case 7
          .TextMatrix(0, c) = "YTD " & strDate
          .ColWidth(c) = cstPeriodDescColW
        Case 12
          .TextMatrix(0, c) = "LTM " & strDate
          .ColWidth(c) = cstPeriodDescColW
      End Select
    Next c
      
    .AddItem "Months"
    .CellAlignment = flexAlignLeftCenter
     
    .AddItem ""
    r = 2
    ' Fill hdr row 2
    For c = 0 To .Cols - 1
    
      .FillStyle = flexFillSingle
      .Row = r
      .col = c
      .CellFontBold = True
      
      Select Case c
        Case 0
          .CellAlignment = flexAlignLeftCenter
          .TextMatrix(r, 0) = "OPERATIONS"
        Case Else
          .CellAlignment = flexAlignCenterCenter
          .CellFontUnderline = True
          m = c Mod (5)
          Select Case m
            Case 1, 4
              .TextMatrix(r, c) = "Last Yr"
            Case 2
              .TextMatrix(r, c) = strTypeDesc '"Current"
            Case 0, 3
              .TextMatrix(r, c) = "Budget"
          End Select
      End Select
    Next c
  End With
End Sub

Sub FillHeaderCap(dbDAO As DAO.Database, ocxFlexGrid As Control, strDate As String)
  Dim c As Integer, m As Integer
  Dim lngRo As Long
  
  With ocxFlexGrid '(R, C)
    .FillStyle = flexFillSingle
    .AddItem ""
    .AddItem ""
    lngRo = .Rows - 1

    ' Fill hdr row 1
    For c = 0 To 13
      .Row = lngRo
      .col = c
      .CellFontBold = True
    
      Select Case c
        Case 0
          .ColAlignment(c) = flexAlignLeftCenter
          .TextMatrix(lngRo, c) = "CAP STRUCTURE"

        Case 4, 5
          .CellAlignment = flexAlignCenterCenter
          .TextMatrix(lngRo, c) = "$ Change"
          
        Case 6
          .CellAlignment = flexAlignLeftCenter
          .CellBackColor = cstAAColor
          .CellForeColor = vbWhite
          .TextMatrix(lngRo, c) = "WRK CAPITAL"
          
          .Row = lngRo - 1
          .col = c
          .CellFontBold = True
          .TextMatrix(lngRo - 1, c) = "EOM " & strDate 'cell above
          .CellAlignment = flexAlignCenterCenter
                      
        Case 7, 11
          .CellAlignment = flexAlignCenterCenter
          .CellFontUnderline = True
          .TextMatrix(lngRo, c) = "Last Yr"
          
        Case 8, 12
          .CellAlignment = flexAlignCenterCenter
          .CellFontUnderline = True
          .TextMatrix(lngRo, c) = strTypeDesc '"Current"
          
        Case 13
          .CellAlignment = flexAlignCenterCenter
          .CellFontUnderline = True
          .TextMatrix(lngRo, c) = "Budget"
          
        Case 9
          .CellAlignment = flexAlignCenterCenter
          .TextMatrix(lngRo, c) = "Change %"
          
        Case 10
          .CellAlignment = flexAlignLeftCenter
          .CellBackColor = cstAAColor
          .CellForeColor = vbWhite
          .TextMatrix(lngRo, c) = "EST MRKT EV"
     
       End Select
    Next c
    
  End With
End Sub

Function GetSummaryDataSQL(lngCompID As Long) As String
  Dim dbDAO As DAO.Database
  Dim rst As DAO.Recordset
  Dim f As Integer, iLastFld As Integer, iSkippedFlds As Integer
  Dim strFld As String, strSQL As String, strFieldsSQL As String 'to build as per CompFields

  Const cstSELECT_Start = "SELECT tblSummaryFormData.FormCat,tblSummaryFormData.Col," & _
                          "Sum(tblSummaryFormData.Months) AS Months, "
  Const cstSELECT_End = " FROM tblSummaryFormData GROUP BY tblSummaryFormData.FormCat," & _
                        "tblSummaryFormData.Col ORDER BY tblSummaryFormData.Col, " & _
                        "Sum(tblSummaryFormData.Months), tblSummaryFormData.FormCat;"
  Const cstPrefix = "Sum([tblSummaryFormData]!" ' & strFld="[" & fieldname & "]"
  Const cstSuffix = ") AS " ' + strFld + comma
  
  Set dbDAO = CurrentDb
  Set rst = dbDAO.OpenRecordset("tblSummaryFormData")
  rst.MoveLast
  rst.MoveFirst
  iLastFld = rst.Fields.Count - 1
  iSkippedFlds = 7
  'iSkippedFlds= 6 (index) +1 ( to start on next pos)
  ' SummaryFormData table non-currency fields:
  '     FormCat, Col, ForecastID, ForecastType, DefBudget, PeriodEnd, Months
  '     [   0  ,  1 ,     2     ,       3     ,      4   ,   5      ,    6  ]
  For f = iSkippedFlds To iLastFld
    strFld = "[" & rst.Fields(f).Name & "]"
    'If f = iLastFld Then
    strFieldsSQL = strFieldsSQL & cstPrefix & strFld & cstSuffix & strFld
    'Else
     If f <> iLastFld Then strFieldsSQL = strFieldsSQL & ","
    'End If
  Next f
  rst.Close
  Set rst = Nothing
  Set dbDAO = Nothing
  
  strSQL = cstSELECT_Start & strFieldsSQL & cstSELECT_End
  GetSummaryDataSQL = strSQL
  
End Function

Sub FillSummaryGrid(ocxGrid As Control, dteEndDate As Date, strCo As String, sglMultiple As Single)
' Called by RefreshSummaryGrid.
' Does not perform transposition: Uses the transposed table, tblSummaryFinal as rstData source.
'
  Dim dbDAO As DAO.Database
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim varColMonths As Variant
  Dim lngRstRecs As Long
  Dim strDateDisp As String
  
  Dim c As Integer, iDateMonth As Integer
  Dim iRstColNum As Integer ' read from 1st row in rst (Col)
  Dim iRstMonths As Integer '           2nd row        (Months)
  '------------------------------------------------------------------------------------
  'On Error GoTo FillSummaryGridErr
    
  strDateDisp = Format(dteEndDate, "mmm 'yy")
  iDateMonth = Month(dteEndDate)
  
  Call FillHeaderTop(dbDAO, ocxGrid, strDateDisp)

  Set dbDAO = CurrentDb
  
  ' Get the Col Num & Months Count array:
  Set qdf = dbDAO.QueryDefs("qrySF-Fill-ColMonths")
  Set rst = qdf.OpenRecordset
  rst.MoveLast
  rst.MoveFirst
  lngRstRecs = rst.RecordCount
  If lngRstRecs <> 2 Then 'error
    MsgBox "Error retrieving the ColNum-MonthsCount array (rec count<>2).", vbCritical, _
           cstMDL & ": FillSummaryGrid"
    rst.Close
    qdf.Close
    GoTo FillSummaryGridExit
  End If
  varColMonths = rst.GetRows(lngRstRecs)
  rst.Close
  qdf.Close
  Set rst = Nothing
  Set qdf = Nothing
  
  'Fill grid  second row with rst second row (Months count); idx=1
  For c = cstOffsetCol To UBound(varColMonths, 1)   '4=col offset(1st 3: names, order1, order2, order3)
    iRstColNum = varColMonths(c, 0)
    iRstMonths = varColMonths(c, 1)
    ocxGrid.TextMatrix(1, iRstColNum) = iRstMonths
    ocxGrid.RowHeight(1) = cstRowHeight
  Next c
  ocxGrid.ColWidth(15) = cstChangeColW
  
  Call FillSectionOps(dbDAO, ocxGrid, varColMonths, iDateMonth)
  
  Call FillHeaderCap(dbDAO, ocxGrid, strDateDisp)
  
  Call FillSectionDebt(dbDAO, ocxGrid, varColMonths)
  
  Call FillSectionEq(dbDAO, ocxGrid, varColMonths)
  
  Call FillSectionWC(dbDAO, ocxGrid, varColMonths)
  
  Call FillFinancialRatios(dbDAO, ocxGrid, sglMultiple)


FillSummaryGridExit:
  Set rst = Nothing
  Set qdf = Nothing
  Set dbDAO = Nothing
  Erase varColMonths
  Exit Sub

FillSummaryGridErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , cstMDL & ": FillSummaryGrid"
  Resume FillSummaryGridExit

End Sub
 
Sub FillSectionOps(dbDAO As DAO.Database, ocxGrid As Control, varColMon As Variant, iMon As Integer)
'
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim varVal As Variant, varData As Variant, var1 As Variant, var2 As Variant
  
  Dim lngRstRecs As Long, lngRow As Long
  Dim r As Integer, c As Integer, d1 As Integer, d2 As Integer
  Dim iRstColNum As Integer, iRstMonths As Integer
  
  Dim lngType As Long, lngPriority As Long
  Dim lngAmt As Long, lngDiff As Long
  '
  Dim lngRevM3 As Long
  Dim lngRevYTD1 As Long, lngRevYTD2 As Long, lngRevYTD3 As Long
  Dim lngRevLTM1 As Long, lngRevLTM2 As Long, lngRevLTM3 As Long
  
  Dim lngGProfM3 As Long
  Dim lngGProfYTD1 As Long, lngGProfYTD2 As Long, lngGProfYTD3 As Long
  Dim lngGProfLTM1 As Long, lngGProfLTM2 As Long, lngGProfLTM3 As Long

  Dim lngEbitdaM3 As Long, lngAdjEbitdaM3 As Long
  Dim lngEbitdaYTD1 As Long, lngEbitdaYTD2 As Long, lngEbitdaYTD3 As Long
  Dim lngAdjEbitdaYTD1 As Long, lngAdjEbitdaYTD2 As Long, lngAdjEbitdaYTD3 As Long
  
  'For col resizing:
  Dim lngColW As Long, lngW As Long
  Dim str As String

  ' variables for annualization:
  Dim blnAnnualized As Boolean
  Const cstAnnualized = "NOTE: Operations amounts of incomplete series have been annualized."
  '------------------------------------------------------------------------------
  blnAnnualized = False
  iRstColNum = 0: iRstMonths = 0: 'lngColW = 0
  '------------------------------------------------------------------------------

  ' *** Operations & Margins section:
  Set qdf = dbDAO.QueryDefs("qrySF-Fill-Ops")
  Set rst = qdf.OpenRecordset
  rst.MoveLast
  rst.MoveFirst
  lngRstRecs = rst.RecordCount
  varData = rst.GetRows(lngRstRecs)
  rst.Close
  qdf.Close
  Set rst = Nothing
  Set qdf = Nothing
  
  With ocxGrid
    .FillStyle = flexFillSingle

    For r = 0 To UBound(varData, 2)
      
      lngType = varData(1, r)
      lngPriority = varData(3, r)
      .AddItem ""
      lngRow = .Rows - 1
      str = varData(0, r) 'fld name
      .TextMatrix(lngRow, 0) = str
      .RowHeight(lngRow) = cstRowHeight
       
      lngW = (Len(str) * cstTwips)
      If lngColW < lngW Then lngColW = lngW 'save largest to resize col 0 on exit
'     SummaryFormData table non-currency fields:
'        FormCat, Col, ForecastID, ForecastType, DefBudget, PeriodEnd, Months
'        [   0  ,  1 ,     2     ,       3     ,      4   ,   5      ,    6  ]
      
      'Debug.Print "lngRow: " & lngRow & "; lngType: " & lngType & "; lngPriority: " & lngPriority

      For c = cstOffsetCol To UBound(varData, 1)
        '4=col offset(1st 3: names, order1, order2, order3)
        varVal = 0
        iRstColNum = varColMon(c, 0)
        iRstMonths = varColMon(c, 1)
        'Save value
        varVal = Nz(varData(c, r), 0)
        
       ' Debug.Print "; iRstColNum: " & iRstColNum & "; varVal: " & varVal
        
        Select Case iRstColNum
          Case 1  'last yr M
            If lngType = 2 Then
              lngIntExpM1 = varVal
            ElseIf lngType <> 5 Then
              Select Case lngPriority
                Case 1
                  lngRevM1 = varVal
                Case 2
                  lngGProfM1 = varVal   'Case 3=operating profit
                Case 4
                  lngEbitdaM1 = varVal
                Case 9
                  lngAdjEbitdaM1 = varVal
              End Select
           End If
          Case 2  'current M
            If lngType = 2 Then
              lngIntExpM2 = varVal
            ElseIf lngType <> 5 Then
              Select Case lngPriority
                Case 1
                  lngRevM2 = varVal
                Case 2
                  lngGProfM2 = varVal
                Case 4
                  lngEbitdaM2 = varVal
                Case 9
                  lngAdjEbitdaM2 = varVal
              End Select
            End If
          Case 3  'budget M
            If lngType <> 2 And lngType <> 5 Then
              Select Case lngPriority
                Case 1
                  lngRevM3 = varVal
                Case 2
                  lngGProfM3 = varVal
                Case 4
                  lngEbitdaM3 = varVal
                Case 9
                  lngAdjEbitdaM3 = varVal
              End Select
            End If
            
          Case 6 'Last yr YTD
            If iRstMonths <> iMon Then  ' same for cols 6, 7, 8
              blnAnnualized = True
              varVal = varVal * (iMon / iRstMonths)
            End If
            If lngType <> 2 And lngType <> 5 Then
              Select Case lngPriority
                Case 1
                  lngRevYTD1 = varVal
                Case 2
                  lngGProfYTD1 = varVal
                Case 4
                  lngEbitdaYTD1 = varVal
                Case 9
                  lngAdjEbitdaYTD1 = varVal
              End Select
            End If
            
          Case 7  'current YTD
            If iRstMonths <> iMon Then  ' same for cols 6, 7, 8
              blnAnnualized = True
              varVal = varVal * (iMon / iRstMonths)
            End If
            If lngType <> 2 And lngType <> 5 Then
              Select Case lngPriority
                Case 1
                  lngRevYTD2 = varVal
                Case 2
                  lngGProfYTD2 = varVal
                Case 4
                  lngEbitdaYTD2 = varVal
                Case 9
                  lngAdjEbitdaYTD2 = varVal
              End Select
            End If
            
          Case 8  'budget YTD
            If iRstMonths <> iMon Then  ' same for cols 6, 7, 8
              blnAnnualized = True
              varVal = varVal * (iMon / iRstMonths)
            End If
            If lngType <> 2 And lngType <> 5 Then
              Select Case lngPriority
                Case 1
                  lngRevYTD3 = varVal
                Case 2
                  lngGProfYTD3 = varVal
                Case 4
                  lngEbitdaYTD3 = varVal
                Case 9
                  lngAdjEbitdaYTD3 = varVal
              End Select
            End If
            
          Case 11 'last yr LTM
            If iRstMonths <> 12 Then ' same for cols 11, 12, 13
              blnAnnualized = True
              varVal = varVal * (12 / iRstMonths)
            End If
            If lngType = 2 Then
              lngIntExpLTM1 = varVal
            ElseIf lngType <> 5 Then
              Select Case lngPriority
                Case 1
                  lngRevLTM1 = varVal
                Case 2
                  lngGProfLTM1 = varVal
                Case 4
                  lngEbitdaLTM1 = varVal
                Case 9
                  lngAdjEbitdaLTM1 = varVal
              End Select
            End If
            
          Case 12 'current LTM
            If iRstMonths <> 12 Then ' same for cols 11, 12, 13
              blnAnnualized = True
              varVal = varVal * (12 / iRstMonths)
            End If
            If lngType = 2 Then
              lngIntExpLTM2 = varVal
            ElseIf lngType <> 5 Then
              Select Case lngPriority
                Case 1
                  lngRevLTM2 = varVal
                Case 2
                  lngGProfLTM2 = varVal
                Case 4
                  lngEbitdaLTM2 = varVal
                Case 9
                  lngAdjEbitdaLTM2 = varVal
              End Select
            End If
            
          Case 13 'budget LTM
            If iRstMonths <> 12 Then ' same for cols 11, 12, 13
              blnAnnualized = True
              varVal = varVal * (12 / iRstMonths)
            End If
            If lngType = 2 Then
              lngIntExpLTM3 = varVal
            ElseIf lngType <> 5 Then
              Select Case lngPriority
                Case 1
                  lngRevLTM3 = varVal
                Case 2
                  lngGProfLTM3 = varVal
                Case 4
                  lngEbitdaLTM3 = varVal
                Case 9
                  lngAdjEbitdaLTM3 = varVal
              End Select
            End If
        End Select
        .CellAlignment = flexAlignRightCenter
        .TextMatrix(lngRow, iRstColNum) = Format(varVal, cstCurrDisp)
      Next c
    Next r
    Erase varData
'Debug.Print "Section Ops end"
'Debug.Print "lngEbitdaLTM1: " & lngEbitdaLTM1, "lngEbitdaLTM2: " & lngEbitdaLTM2, "lngEbitdaLTM3: " & lngEbitdaLTM3
    
    .ColWidth(0) = lngColW                     ' Resize firstData column
    'Fill margins hdr rows
    Call FillAddBlankRow(ocxGrid, 1, 15)
    lngRow = .Rows - 1
    .Row = lngRow
    .col = 0
    .CellAlignment = flexAlignRightCenter
    .CellFontBold = True
    .TextMatrix(lngRow, 0) = "Margins"
      
    'Fill gross margin row
    .AddItem ""
    lngRow = .Rows - 1
    .Row = lngRow
    .col = 0
    .CellAlignment = flexAlignRightCenter
    .CellFontBold = True
    .TextMatrix(lngRow, 0) = "Gross Margin"
  
    For c = 1 To 15
      varVal = 0: var1 = 0: var2 = 0: d1 = 0: d2 = 0 'reset
      Select Case c
        Case 1
          If lngRevM1 <> 0 Then varVal = lngGProfM1 / lngRevM1 '* cstCent)
        Case 2
          If lngRevM2 <> 0 Then varVal = lngGProfM2 / lngRevM2
        Case 3
          If lngRevM3 <> 0 Then varVal = lngGProfM3 / lngRevM3
               
        Case 4, 9, 14
          d1 = c - 2
          d2 = c - 3
          var1 = Trim(.TextMatrix(lngRow, d1))
          var1 = Left(var1, Len(var1) - 1)
          var2 = Trim(.TextMatrix(lngRow, d2))
          var2 = Left(var2, Len(var2) - 1)
          varVal = Eval((var1 - var2)) / cstCent
        Case 5, 10, 15
          d1 = c - 3
          d2 = c - 2
          var1 = Trim(.TextMatrix(lngRow, d1))
          var1 = Left(var1, Len(var1) - 1)
          var2 = Trim(.TextMatrix(lngRow, d2))
          var2 = Left(var2, Len(var2) - 1)
          varVal = Eval(var1 - var2) / cstCent
          
        Case 6
          If lngRevYTD1 <> 0 Then varVal = lngGProfYTD1 / lngRevYTD1
        Case 7
          If lngRevYTD2 <> 0 Then varVal = lngGProfYTD2 / lngRevYTD2
        Case 8
          If lngRevYTD3 <> 0 Then varVal = lngGProfYTD3 / lngRevYTD3
        Case 11
          If lngRevLTM1 <> 0 Then varVal = lngGProfLTM1 / lngRevLTM1
        Case 12
          If lngRevLTM2 <> 0 Then varVal = lngGProfLTM2 / lngRevLTM2
        Case 13
          If lngRevLTM3 <> 0 Then varVal = lngGProfLTM3 / lngRevLTM3
       
      End Select
      .TextMatrix(lngRow, c) = Format(varVal, cstChangeDisp) 'varVal * cstCent
    Next c
    
    'Fill ebidta margin row
    .AddItem ""
    lngRow = .Rows - 1
    .Row = lngRow
    .col = 0
    .CellAlignment = flexAlignRightCenter
    .CellFontBold = True
    .TextMatrix(lngRow, 0) = "EBITDA"

    For c = 1 To 15
      varVal = 0: var1 = 0: var2 = 0: d1 = 0: d2 = 0 'reset
      Select Case c
        Case 1
          If lngRevM1 <> 0 Then varVal = (lngEbitdaM1 + lngAdjEbitdaM1) / lngRevM1
        Case 2
          If lngRevM2 <> 0 Then varVal = (lngEbitdaM2 + lngAdjEbitdaM2) / lngRevM2
        Case 3
          If lngRevM3 <> 0 Then varVal = (lngEbitdaM3 + lngAdjEbitdaM3) / lngRevM3
              
        Case 4, 9, 14
          d1 = c - 2
          d2 = c - 3
          var1 = Trim(.TextMatrix(lngRow, d1))
          var1 = Left(var1, Len(var1) - 1)
          var2 = Trim(.TextMatrix(lngRow, d2))
          var2 = Left(var2, Len(var2) - 1)
          varVal = Eval(var1 - var2) / cstCent
                  
        Case 5, 10, 15
          d1 = c - 3
          d2 = c - 2
          var1 = Trim(.TextMatrix(lngRow, d1))
          var1 = Left(var1, Len(var1) - 1)
          var2 = Trim(.TextMatrix(lngRow, d2))
          var2 = Left(var2, Len(var2) - 1)
          varVal = Eval(var1 - var2) / cstCent
        
        Case 6
          If lngRevYTD1 <> 0 Then varVal = (lngEbitdaYTD1 + lngAdjEbitdaYTD1) / lngRevYTD1
        Case 7
          If lngRevYTD2 <> 0 Then varVal = (lngEbitdaYTD2 + lngAdjEbitdaYTD2) / lngRevYTD2
        Case 8
          If lngRevYTD3 <> 0 Then varVal = (lngEbitdaYTD3 + lngAdjEbitdaYTD3) / lngRevYTD3
        Case 11
          If lngRevLTM1 <> 0 Then varVal = (lngEbitdaLTM1 + lngAdjEbitdaLTM1) / lngRevLTM1
        Case 12
          If lngRevLTM2 <> 0 Then varVal = (lngEbitdaLTM2 + lngAdjEbitdaLTM2) / lngRevLTM2
        Case 13
          If lngRevLTM3 <> 0 Then varVal = (lngEbitdaLTM3 + lngAdjEbitdaLTM3) / lngRevLTM3
      End Select
      .TextMatrix(lngRow, c) = Format(varVal, cstChangeDisp)  '
    Next c

    If blnAnnualized Then
      .Parent!lblWarn.Visible = True
      .Parent!lblWarn.Caption = cstAnnualized
    End If
  End With
 
  Call FillChangeColsOps(dbDAO, ocxGrid)
  
  Call FillAddBlankRow(ocxGrid, 1, 15)  'will not merge first cell
  lngRow = ocxGrid.Rows - 1
  ' *** END Operations & Margins section
End Sub

Sub FillSectionDebt(dbDAO As DAO.Database, ocxGrid As Control, varColMon As Variant)
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim varVal As Variant, varData As Variant
  
  Dim lngRstRecs As Long, lngRowCount As Long, lngRow As Long
  Dim r As Integer, c As Integer
  Dim iRstColNum As Integer 'read from 1st row in rst (Col)
  Dim iRstMonths As Integer '          2nd row        (Months)

  Dim lngType As Long, lngPriority As Long
  Dim lngAmt As Long, lngDiff As Long
  'For col resizing:
  Dim lngColW As Long, lngW As Long
  Dim str As String
  
  ' *** Capital Structure (Debt) section:
  Set qdf = dbDAO.QueryDefs("qrySF-Fill-Debt")
  Set rst = qdf.OpenRecordset
  rst.MoveLast
  rst.MoveFirst
  lngRstRecs = rst.RecordCount
  varData = rst.GetRows(lngRstRecs)
  rst.Close
  qdf.Close
  Set rst = Nothing
  Set qdf = Nothing
  lngTotDebtRow = 0: lngTotEqRow = 0
  
  With ocxGrid
    .FillStyle = flexFillSingle
    
    lngSrSubM1 = 0: lngSrSubM2 = 0: lngSrSubM3 = 0
    lngTotDebtM1 = 0: lngTotDebtM2 = 0: lngTotDebtM3 = 0
    
    For r = 0 To UBound(varData, 2)
      lngType = varData(1, r)
      lngPriority = varData(3, r)
      
      .AddItem ""
      lngRow = .Rows - 1
      If r = 0 Then lngFirstCapRow = lngRow
      
      str = varData(0, r) 'fld name
      lngW = (Len(str) * cstTwips)
      If lngColW < lngW Then lngColW = lngW 'save largest to resize col 0 on exit
      .Row = lngRow
      .col = 0
      .TextMatrix(lngRow, 0) = varData(0, r)  'fld name
      .CellAlignment = flexAlignLeftCenter
      .CellFontBold = False
      
      For c = cstOffsetCol To 6   ' up to M3, if there
                              'UBound(varData, 1)  ' from MonthEnd section to LTM current
        varVal = 0
        iRstColNum = varColMon(c, 0)
        varVal = Nz(varData(c, r), 0)
        .Row = lngRow
        .col = iRstColNum
        .CellAlignment = flexAlignRightCenter
        Select Case iRstColNum
          Case 1  'last yr M
            If lngType > 6 Then
              lngTotDebtM1 = lngTotDebtM1 + varVal                     'accumulate Tot Debt
              If lngPriority < 3 Then lngSrSubM1 = lngSrSubM1 + varVal 'accumulate Tot SrSub
            End If
            .TextMatrix(lngRow, iRstColNum) = Format(varVal, cstCurrDisp)
          Case 2  'current M
            If lngType > 6 Then
              lngTotDebtM2 = lngTotDebtM2 + varVal
              If lngPriority < 3 Then lngSrSubM2 = lngSrSubM2 + varVal
            End If
            .TextMatrix(lngRow, iRstColNum) = Format(varVal, cstCurrDisp)
          Case 3 ' budget M
            If lngType > 6 Then
              lngTotDebtM3 = lngTotDebtM3 + varVal
              If lngPriority < 3 Then lngSrSubM3 = lngSrSubM3 + varVal
            End If
            .TextMatrix(lngRow, iRstColNum) = Format(varVal, cstCurrDisp)
        End Select
      Next c
    Next r
    Erase varData
    .ColWidth(0) = lngColW                     ' Resize firstData column
    
    .AddItem ""
    lngRow = .Rows - 1
    .Row = lngRow
    .col = 0
    lngTotDebtRow = lngRow
    .TextMatrix(lngRow, 0) = "Tot Debt"
    .CellAlignment = flexAlignRightCenter

    .CellFontBold = True
    .TextMatrix(lngRow, 1) = Format(lngTotDebtM1, cstCurrDisp)
    .TextMatrix(lngRow, 2) = Format(lngTotDebtM2, cstCurrDisp)
    .TextMatrix(lngRow, 3) = Format(lngTotDebtM3, cstCurrDisp)
    Call FillAddBlankRow(ocxGrid, 0, 5)
  End With
    
End Sub

Sub FillSectionEq(dbDAO As DAO.Database, ocxGrid As Control, varColMon As Variant)
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim varVal As Variant, varData As Variant
  
  Dim lngRstRecs As Long, lngRowCount As Long, lngRow As Long
  Dim r As Integer, c As Integer
  Dim iRstColNum As Integer 'read from 1st row in rst (Col)
  Dim iRstMonths As Integer '          2nd row        (Months)

  Dim lngType As Long, lngPriority As Long
  Dim lngAmt As Long, lngDiff As Long
   
  Dim lngTotEqM1 As Long, lngTotEqM2 As Long, lngTotEqM3 As Long
  
  lngTotEqM1 = 0: lngTotEqM2 = 0: lngTotEqM3 = 0
  lngTotCapM1 = 0: lngTotCapM2 = 0: lngTotCapM3 = 0
  
  ' *** Capital Structure (Equity) section:
  Set qdf = dbDAO.QueryDefs("qrySF-Fill-Eq")
  Set rst = qdf.OpenRecordset
  rst.MoveLast
  rst.MoveFirst
  lngRstRecs = rst.RecordCount
  varData = rst.GetRows(lngRstRecs)
  rst.Close
  qdf.Close
  Set rst = Nothing
  Set qdf = Nothing
  
  With ocxGrid
    .FillStyle = 0

    For r = 0 To UBound(varData, 2)
      lngType = varData(1, r)
      lngPriority = varData(3, r)
      
      .AddItem ""
      lngRow = .Rows - 1
      .TextMatrix(lngRow, 0) = varData(0, r)  'fld name
        
      For c = cstOffsetCol To 6 'UBound(varData, 1)
        varVal = 0
        iRstColNum = varColMon(c, 0)
        varVal = Nz(varData(c, r), 0)
        
        Select Case iRstColNum
          Case 1  'last yr M
            lngTotEqM1 = lngTotEqM1 + varVal
            .TextMatrix(lngRow, iRstColNum) = Format(varVal, cstCurrDisp)
          Case 2  'current M
            lngTotEqM2 = lngTotEqM2 + varVal
            .TextMatrix(lngRow, iRstColNum) = Format(varVal, cstCurrDisp)
          Case 3 'budget M
            lngTotEqM3 = lngTotEqM3 + varVal
            .TextMatrix(lngRow, iRstColNum) = Format(varVal, cstCurrDisp)
        End Select
      Next c
    Next r
    Erase varData
  
    .AddItem ""
    lngRow = .Rows - 1
    .Row = lngRow
    .col = 0
    .CellAlignment = flexAlignRightCenter
    .CellFontBold = True
    lngTotEqRow = lngRow
    .TextMatrix(lngRow, 0) = "Tot Equity"
    .TextMatrix(lngRow, 1) = Format(lngTotEqM1, cstCurrDisp)
    .TextMatrix(lngRow, 2) = Format(lngTotEqM2, cstCurrDisp)
    .TextMatrix(lngRow, 3) = Format(lngTotEqM3, cstCurrDisp)
    Call FillAddBlankRow(ocxGrid, 0, 5)
  
    .AddItem ""
    lngRow = .Rows - 1
    lngLastCapRow = lngRow
    .Row = lngRow
    .col = 0
    .CellAlignment = flexAlignRightCenter
    .CellFontBold = True
    .TextMatrix(lngRow, 0) = "Tot Capitalization"
    lngTotCapM1 = lngTotDebtM1 + lngTotEqM1
    lngTotCapM2 = lngTotDebtM2 + lngTotEqM2
    lngTotCapM3 = lngTotDebtM3 + lngTotEqM3
    .TextMatrix(lngRow, 1) = Format(lngTotCapM1, cstCurrDisp)
    .TextMatrix(lngRow, 2) = Format(lngTotCapM2, cstCurrDisp)
    .TextMatrix(lngRow, 3) = Format(lngTotCapM3, cstCurrDisp)
    
    Call FillChangeColsCap(ocxGrid, lngFirstCapRow, lngLastCapRow)
    ' *** END Capital Structure section
  End With
  
End Sub

Sub FillSectionWC(dbDAO As DAO.Database, ocxGrid As Control, varColMon As Variant)
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim varVal As Variant, varData As Variant
  
  Dim lngRstRecs As Long, lngRowCount As Long, lngRow As Long, lngWCRatiosRow As Long
  Dim r As Integer, c As Integer
  Dim iRstColNum As Integer  'read from 1st row in rst (Col)
  Dim iRstMonths1 As Integer '          2nd row        (Months)
  Dim iRstMonths2 As Integer
  
  Dim lngType As Long, lngPriority As Long
  Dim lngAmt As Long, lngDiff As Long
  Dim lngAR1 As Long, lngAR2 As Long, lngAP1 As Long, lngAP2 As Long
  Dim lngInvent1 As Long, lngInvent2 As Long
  Dim lngRevMax1 As Long, lngRevMax2 As Long
  Dim lngBorBase1 As Long, lngBorBase2 As Long
  Dim lngExtrapSales1 As Long, lngExtrapSales2 As Long
  Dim lngExtrapCOG1 As Long, lngExtrapCOG2 As Long
  Dim blnMaxRev As Boolean, blnSkipWCPartI As Boolean, blnSkipWCPartII As Boolean

  ' *** Working-Cap section: (includes ratio fields)
  Set qdf = dbDAO.QueryDefs("qrySF-Fill-WorkCap")
  Set rst = qdf.OpenRecordset
  If rst.AbsolutePosition <> -1 Then
    rst.MoveLast
    
    ' retrieve last row if revmax
    If rst.Fields(3) = 9 Then 'there is a Revolver Max amt in one or more columns
      blnMaxRev = True
      If varColMon(cstOffsetCol, 0) = 1 Then
        lngRevMax1 = rst.Fields(cstOffsetCol).Value     'EOM-LastYr Revolver Max value
      ElseIf varColMon(cstOffsetCol + 1, 0) = 2 Then
        lngRevMax2 = rst.Fields(cstOffsetCol + 1).Value 'EOM-Current Revolver Max value
      End If
    End If
    rst.MoveFirst
    lngRstRecs = rst.RecordCount
    If blnMaxRev Then lngRstRecs = lngRstRecs - 1 'remove last row from array
    If lngRstRecs > 0 Then varData = rst.GetRows(lngRstRecs)
  End If
  rst.Close
  qdf.Close
  Set rst = Nothing
  Set qdf = Nothing
    
  blnSkipWCPartI = IsNull(varData)
  'Note: The code to conditinally fill the "Inventory" section was created
  '      prior to setting the Inventory field as a required field.
  '      Even though varData will always at least contain the Inventory field
  '      (regardless it has data or not), thus setting blnSkipWCPartI as False,
  '      it is kept until a decision is made to fill that section, even though
  '      it may have no data.
  '
  lngRow = lngFirstCapRow
  
  With ocxGrid
    .FillStyle = 0
    .col = 6
    
    If Not blnSkipWCPartI Then
      lngRowCount = -1
  
      For r = 0 To UBound(varData, 2)
        lngRowCount = lngRow + r
                
        lngType = varData(1, r)
        lngPriority = varData(3, r)
           
        .Row = lngRowCount
        .CellAlignment = flexAlignRightCenter
        .TextMatrix(lngRowCount, 6) = varData(0, r)  'fld name
        
        For c = cstOffsetCol To 5 'MonthEnd section only: Last Yr, Current (if there)
          varVal = 0
          iRstColNum = varColMon(c, 0)
        
          varVal = Nz(varData(c, r), 0)
          .TextMatrix(lngRowCount, (iRstColNum + 6)) = Format(varVal, cstCurrDisp)
          
          Select Case iRstColNum
            Case 1  'last yr M
              iRstMonths1 = varColMon(c, 1)
              If iRstMonths1 = 0 Then iRstMonths1 = .TextMatrix(1, 2) ' Months cell
              Select Case lngType
                Case 3
                  If lngPriority = 1 Then
                    lngAR1 = varVal
                  Else
                    lngInvent1 = varVal
                  End If
                  
                Case Else
                  lngAP1 = varVal
               End Select
      
            Case 2  'current M
              iRstMonths2 = varColMon(c, 1)
              If iRstMonths2 = 0 Then iRstMonths2 = .TextMatrix(1, 3) ' Months cell
              Select Case lngType
                Case 3
                  If lngPriority = 1 Then
                    lngAR2 = varVal
                  Else
                    lngInvent2 = varVal
                  End If
                  
                Case Else
                  lngAP2 = varVal
               End Select
          End Select
        Next c
      Next r
      Erase varData
      
      If lngRowCount > 0 Then lngRow = lngRowCount
      lngRow = lngRow + 1
 
      .Row = lngRow
      .TextMatrix(lngRow, 6) = "Net W/C: "
      .CellAlignment = flexAlignRightCenter
      .TextMatrix(lngRow, 7) = Format((lngAR1 + lngInvent1 - lngAP1), cstCurrDisp)
      .TextMatrix(lngRow, 8) = Format((lngAR2 + lngInvent2 - lngAP2), cstCurrDisp)
      
      Call FillChangeColsWC(dbDAO, ocxGrid, lngFirstCapRow, lngRow)
      
      lngRow = lngRow + 1
      Call FillAddBlankRow(ocxGrid, 6, 9)
    End If  ' Not blnSkipWCPartI
      
    If (lngAR1 + lngInvent1 + lngAP1 + lngRevMax1 + _
        lngAR2 + lngInvent2 + lngAP2 + lngRevMax2) = 0 Then blnSkipWCPartII = True
        '(All zero-valued: no need to write the W/C ratios section)
    
    If Not blnSkipWCPartII Then
      lngRow = lngRow + 1
      If lngLastCapRow < lngRow Then
        .AddItem ""
        lngRow = .Rows - 1
      End If
      
      .Row = lngRow
      .CellFontBold = True
      .CellBackColor = cstAAColor
      .CellForeColor = vbWhite
      .TextMatrix(lngRow, 6) = "W/C RATIOS"
      .CellAlignment = flexAlignLeftCenter
      
      lngRow = lngRow + 1
      lngWCRatiosRow = lngRow
      
      If lngLastCapRow < lngRow Then
        .AddItem ""
        lngRow = .Rows - 1
      End If
            
      .Row = lngRow
      .CellAlignment = flexAlignRightCenter
      .TextMatrix(lngRow, 6) = "Max Revolver"
      .CellFontBold = False

      .TextMatrix(lngRow, 7) = Format(lngRevMax1, cstCurrDisp)
      .TextMatrix(lngRow, 8) = Format(lngRevMax2, cstCurrDisp)
      
      If Not blnSkipWCPartI Then
        lngRow = lngRow + 1 ' .Rows - 1
        If lngLastCapRow < lngRow Then
          .AddItem ""
          lngRow = .Rows - 1
        End If
        
        .Row = lngRow
        .CellAlignment = flexAlignRightCenter
        .TextMatrix(lngRow, 6) = "Est Bor'ing Base"
        .CellFontBold = False
        lngBorBase1 = (0.85 * lngAR1) + (0.5 * lngInvent1)
        lngBorBase2 = (0.85 * lngAR2) + (0.5 * lngInvent2)
        .TextMatrix(lngRow, 7) = Format(lngBorBase1, cstCurrDisp)
        .TextMatrix(lngRow, 8) = Format(lngBorBase2, cstCurrDisp)
        
        lngRow = lngRow + 1
        If lngLastCapRow < lngRow Then
          .AddItem ""
          lngRow = .Rows - 1
        End If
        
        .Row = lngRow
        .CellAlignment = flexAlignRightCenter
        .TextMatrix(lngRow, 6) = "Est Availability"
        .CellFontBold = False

        .TextMatrix(lngRow, 7) = Format(lngRevMax1 - lngBorBase1, cstCurrDisp)
        .TextMatrix(lngRow, 8) = Format(lngRevMax2 - lngBorBase2, cstCurrDisp)
        
        lngRow = lngRow + 1
        If lngLastCapRow < lngRow Then
          .AddItem ""
          lngRow = .Rows - 1
        End If
        
        'Extrapolate Sales:
        If iRstMonths1 = 0 Then iRstMonths1 = .TextMatrix(1, 2)
        lngExtrapSales1 = (lngRevM1 * 12 / iRstMonths1)
        If iRstMonths2 = 0 Then iRstMonths2 = .TextMatrix(1, 3)
        lngExtrapSales2 = (lngRevM2 * 12 / iRstMonths2)
        
        .Row = lngRow
        .CellAlignment = flexAlignRightCenter
        .TextMatrix(lngRow, 6) = "A/R Days"
        .CellFontBold = False
        
        If lngRevM1 <> 0 Then
          .TextMatrix(lngRow, 7) = Format(((lngAR1 * 365) / lngExtrapSales1), "##.#")
        End If
        If lngRevM2 <> 0 Then
          .TextMatrix(lngRow, 8) = Format(((lngAR2 * 365) / lngExtrapSales2), "##.#")
        End If
        
        lngRow = lngRow + 1
        If lngLastCapRow < lngRow Then
          .AddItem ""
          lngRow = .Rows - 1
        End If
        
        'Extrapolate COG:
        lngExtrapCOG1 = (lngRevM1 - lngGProfM1) * (12 / iRstMonths1)
        lngExtrapCOG2 = (lngRevM2 - lngGProfM2) * (12 / iRstMonths2)
        
        .Row = lngRow
        .CellAlignment = flexAlignRightCenter
        .TextMatrix(lngRow, 6) = "Inventory Turns"
        
        If lngInvent1 <> 0 Then
          .TextMatrix(lngRow, 7) = Format((lngExtrapCOG1 / lngInvent1), "#.#x")
        End If
        If lngInvent2 <> 0 Then
          .TextMatrix(lngRow, 8) = Format((lngExtrapCOG2 / lngInvent2), "#.#x")
        End If
        
        lngRow = lngRow + 1
        If lngLastCapRow < lngRow Then
          .AddItem ""
          lngRow = .Rows - 1
        End If
       
        .Row = lngRow
        .CellAlignment = flexAlignRightCenter
        .TextMatrix(lngRow, 6) = "A/P Days"

        If lngExtrapCOG1 <> 0 Then
          .TextMatrix(lngRow, 7) = Format((lngAP1 * 365) / lngExtrapCOG1, "##.#")
        End If
        If lngExtrapCOG2 <> 0 Then
          .TextMatrix(lngRow, 8) = Format((lngAP2 * 365) / lngExtrapCOG2, "##.#")
        End If
      '/////////////////////////////////////////////////////////////////////////
       Call FillChangeColsWC(dbDAO, ocxGrid, lngWCRatiosRow, lngRow)
      End If 'not blnSkipWCPartI
    End If ' not blnSkipWCPartII
    lngLastWCRow = lngRow
  End With
  
End Sub

Sub FillFinancialRatios(dbDAO As DAO.Database, ocxGrid As Control, sglMulti As Single)
  Dim lngRow As Long, lngLastRow As Long
  Dim lngNetEBITDALTM1 As Long, lngNetEBITDALTM2 As Long, lngNetEBITDALTM3 As Long
  Dim lngMEV1 As Long, lngMEV2 As Long, lngMEV3 As Long
  
  Const cstRatioDisp = "#.#x"
  Const cstNewTwips = 100
  
  If lngLastCapRow < lngLastWCRow Then 'take the largest
    lngLastRow = lngLastWCRow
  Else
    lngLastRow = lngLastCapRow
  End If
  lngRow = lngFirstCapRow 'ini
  
  ' Calc
  lngNetEBITDALTM1 = lngEbitdaLTM1 + lngAdjEbitdaLTM1
  lngNetEBITDALTM2 = lngEbitdaLTM2 + lngAdjEbitdaLTM2
  lngNetEBITDALTM3 = lngEbitdaLTM3 + lngAdjEbitdaLTM3
  lngMEV1 = lngEbitdaLTM1 * sglMulti
  lngMEV2 = lngEbitdaLTM2 * sglMulti
  lngMEV3 = lngEbitdaLTM3 * sglMulti
  
  
  With ocxGrid
    ' Because the capital structure section has at least 9 rows (with totals)
    ' there is no need to check if a new row is needed until the tenth row
    .col = 10
    
    .Row = lngRow
    .CellAlignment = flexAlignRightCenter
    ' ** Mrkt val section:
    .TextMatrix(lngRow, 10) = "EBITDA"
    .TextMatrix(lngRow, 11) = Format(lngNetEBITDALTM1, cstCurrDisp)
    .TextMatrix(lngRow, 12) = Format(lngNetEBITDALTM2, cstCurrDisp)
    .TextMatrix(lngRow, 13) = Format(lngNetEBITDALTM3, cstCurrDisp)
  
    lngRow = lngRow + 1
    .Row = lngRow
    .CellAlignment = flexAlignRightCenter
    .TextMatrix(lngRow, 10) = "Multiple"
    .TextMatrix(lngRow, 11) = Format(sglMulti, "#.#0x")
    .TextMatrix(lngRow, 12) = Format(sglMulti, "#.#0x")
    .TextMatrix(lngRow, 13) = Format(sglMulti, "#.#0x")
    
    lngRow = lngRow + 1
    .Row = lngRow
    .CellAlignment = flexAlignRightCenter
    .TextMatrix(lngRow, 10) = "Ent. Val"
    .TextMatrix(lngRow, 11) = Format(lngMEV1, cstCurrDisp)
    .TextMatrix(lngRow, 12) = Format(lngMEV2, cstCurrDisp)
    .TextMatrix(lngRow, 13) = Format(lngMEV3, cstCurrDisp)
    
    lngRow = lngRow + 1
    .Row = lngRow
    .CellAlignment = flexAlignRightCenter
    .TextMatrix(lngRow, 10) = "- Net Debt"
    .TextMatrix(lngRow, 11) = Format(lngTotDebtM1, cstCurrDisp)
    .TextMatrix(lngRow, 12) = Format(lngTotDebtM2, cstCurrDisp)
    .TextMatrix(lngRow, 13) = Format(lngTotDebtM3, cstCurrDisp)

    lngRow = lngRow + 1
    .Row = lngRow
    .CellAlignment = flexAlignRightCenter
    .TextMatrix(lngRow, 10) = "Est Mrkt Eq Val"
    lngMEV1 = lngMEV1 - lngTotDebtM1
    .TextMatrix(lngRow, 11) = Format(lngMEV1, cstCurrDisp)
    lngMEV2 = lngMEV2 - lngTotDebtM2
    .TextMatrix(lngRow, 12) = Format(lngMEV2, cstCurrDisp)
    lngMEV3 = lngMEV3 - lngTotDebtM3
    .TextMatrix(lngRow, 13) = Format(lngMEV3, cstCurrDisp)
    
    lngRow = lngRow + 2   '7th row
    .Row = lngRow
    .CellFontBold = True
    .CellBackColor = cstAAColor
    .CellForeColor = vbWhite
    .CellAlignment = flexAlignLeftCenter
    .TextMatrix(lngRow, 10) = "FIN'L RATIOS"
    
    lngRow = lngRow + 1  '8th row
    .Row = lngRow
    .CellAlignment = flexAlignRightCenter
    .TextMatrix(lngRow, 10) = "EBITDA/Int"

    If lngIntExpLTM1 <> 0 Then
      .TextMatrix(lngRow, 11) = Format((lngNetEBITDALTM1 / lngIntExpLTM1), cstRatioDisp)
    End If
    If lngIntExpLTM2 <> 0 Then
      .TextMatrix(lngRow, 12) = Format((lngNetEBITDALTM2 / lngIntExpLTM2), cstRatioDisp)
    End If
    If lngIntExpLTM3 <> 0 Then
      .TextMatrix(lngRow, 13) = Format((lngNetEBITDALTM3 / lngIntExpLTM3), cstRatioDisp)
    End If
    
    lngRow = lngRow + 1 ' 9th row
    .Row = lngRow
    .CellAlignment = flexAlignRightCenter
    .TextMatrix(lngRow, 10) = "SrDebt/EBITDA"
    If lngNetEBITDALTM1 <> 0 Then
      .TextMatrix(lngRow, 11) = Format((lngSrSubM1 / lngNetEBITDALTM1), cstRatioDisp)
    End If
    If lngNetEBITDALTM2 <> 0 Then
      .TextMatrix(lngRow, 12) = Format((lngSrSubM2 / lngNetEBITDALTM2), cstRatioDisp)
    End If
    If lngNetEBITDALTM3 <> 0 Then
      .TextMatrix(lngRow, 13) = Format((lngSrSubM3 / lngNetEBITDALTM3), cstRatioDisp)
    End If
    
    lngRow = lngRow + 1 ' 10th row
    If lngLastRow < lngRow Then
      .AddItem " "
      lngRow = .Rows - 1
    End If
    .Row = lngRow
    .CellAlignment = flexAlignRightCenter
    .TextMatrix(lngRow, 10) = "TotDebt/EBITDA"
    If lngNetEBITDALTM1 <> 0 Then
      .TextMatrix(lngRow, 11) = Format((lngTotDebtM1 / lngNetEBITDALTM1), cstRatioDisp)
    End If
    If lngNetEBITDALTM2 <> 0 Then
      .TextMatrix(lngRow, 12) = Format((lngTotDebtM2 / lngNetEBITDALTM2), cstRatioDisp)
    End If
    If lngNetEBITDALTM3 <> 0 Then
      .TextMatrix(lngRow, 13) = Format((lngTotDebtM3 / lngNetEBITDALTM3), cstRatioDisp)
    End If
            
    lngRow = lngRow + 1 ' 11th
    If lngLastRow < lngRow Then
      .AddItem " "
      lngRow = .Rows - 1
    End If
    .Row = lngRow
    .CellAlignment = flexAlignRightCenter
    .TextMatrix(lngRow, 10) = "TotDebt/Cap"
    If lngTotCapM1 <> 0 Then
      .TextMatrix(lngRow, 11) = Format((lngTotDebtM1 / lngTotCapM1), "#.#%")
    End If
    If lngTotCapM2 <> 0 Then
      .TextMatrix(lngRow, 12) = Format((lngTotDebtM2 / lngTotCapM2), "#.#%")
    End If
    If lngTotCapM3 <> 0 Then
      .TextMatrix(lngRow, 13) = Format((lngTotDebtM3 / lngTotCapM3), "#.#%")
    End If

    lngRow = lngRow + 1 ' 14th
    If lngLastRow < lngRow Then
      .AddItem " "
      lngRow = .Rows - 1
    End If
    .Row = lngRow
    .CellFontBold = True
    .CellBackColor = cstAAColor
    .CellForeColor = vbWhite
    .CellAlignment = flexAlignRightCenter
    .TextMatrix(lngRow, 10) = "TOT DEBT"
    .TextMatrix(lngRow, 11) = Format(lngTotDebtM1, cstCurrDisp)
    .TextMatrix(lngRow, 12) = Format(lngTotDebtM2, cstCurrDisp)
    .TextMatrix(lngRow, 13) = Format(lngTotDebtM3, cstCurrDisp)
    
  End With
End Sub

Sub FillChangeColsOps(dbDAO As DAO.Database, ocxGrid As Control)
  Dim c As Integer, r As Integer
  Dim lngAmt As Long, lngDiff As Long
  Dim lngLastRow As Long
  
  With ocxGrid

    lngLastRow = .Rows - 1
    For r = 3 To lngLastRow - 3 'stop before the margins row
      For c = cstOffsetCol To 15
        lngAmt = 0
        lngDiff = 0
        .CellAlignment = flexAlignRightCenter
        
        Select Case c
          Case 4, 9, 14
           If IsNumeric(.TextMatrix(r, c - 3)) Then
              lngAmt = CLng(.TextMatrix(r, c - 3))
            End If
            If IsNumeric(.TextMatrix(r, c - 2)) Then
              lngDiff = CLng(.TextMatrix(r, c - 2)) - lngAmt
            End If
            If lngAmt <> 0 Then
             .TextMatrix(r, c) = Format(lngDiff / lngAmt, cstChangeDisp)
            Else
             .TextMatrix(r, c) = Format(lngAmt, cstChangeDisp)
            End If
            
          Case 5, 10, 15
            If IsNumeric(.TextMatrix(r, c - 2)) Then lngAmt = CLng(.TextMatrix(r, c - 2))
            If IsNumeric(.TextMatrix(r, c - 3)) Then
              lngDiff = CLng(.TextMatrix(r, c - 3)) - lngAmt
            End If
            If lngAmt <> 0 Then
             .TextMatrix(r, c) = Format((lngDiff / lngAmt), cstChangeDisp) ' * cstCent
            Else
             .TextMatrix(r, c) = Format(lngAmt, cstChangeDisp)
            End If
        End Select
      Next c
    Next r
  End With
  
End Sub

Sub FillChangeColsCap(ocxGrid As Control, lngStartRow As Long, lngEndRow As Long)
  Dim lngRo As Long
  Dim lngAmt As Long, lngDiff As Long
  Dim c As Long
  
  Const cstLYcol = 4
  Const cstBUcol = 5
  
  With ocxGrid
    For lngRo = lngStartRow To lngEndRow
      ' Skip penultimate blank lines:
      If lngRo = lngTotDebtRow + 1 Or lngRo = lngTotEqRow + 1 Then
        lngRo = lngRo + 1
      End If
      For c = cstLYcol To cstBUcol
        lngAmt = 0
        lngDiff = 0
        Select Case c
        
          Case cstLYcol 'chge from lst yr
            If IsNumeric(.TextMatrix(lngRo, cstLYcol - 3)) Then
              lngAmt = CLng(.TextMatrix(lngRo, cstLYcol - 3))
            End If
            If IsNumeric(.TextMatrix(lngRo, cstLYcol - 2)) Then
              lngDiff = CLng(.TextMatrix(lngRo, cstLYcol - 2))
            End If
            
          Case cstBUcol 'chge from budget
            If IsNumeric(.TextMatrix(lngRo, cstBUcol - 3)) Then
              lngDiff = CLng(.TextMatrix(lngRo, cstBUcol - 3))
            End If
            If IsNumeric(.TextMatrix(lngRo, cstBUcol - 2)) Then
              lngAmt = CLng(.TextMatrix(lngRo, cstBUcol - 2))
            End If
           
        End Select
         lngDiff = lngDiff - lngAmt
        .TextMatrix(lngRo, c) = Format(lngDiff, cstCurrDisp)
      Next c
    Next lngRo
  End With
  
End Sub

Sub FillChangeColsWC(dbDAO As DAO.Database, ocxGrid As Control, lngStartRow As Long, lngEndRow As Long)
  Dim lngRo As Long
  Dim lngAmt As Long, lngDiff As Long
  Const cstCol = 9
  
  With ocxGrid

    For lngRo = lngStartRow To lngEndRow
      lngAmt = 0: lngDiff = 0
      If IsNumeric(.TextMatrix(lngRo, cstCol - 2)) Then lngAmt = CLng(.TextMatrix(lngRo, cstCol - 2))
      If IsNumeric(.TextMatrix(lngRo, cstCol - 1)) Then
        lngDiff = CLng(.TextMatrix(lngRo, cstCol - 1) - lngAmt)
      End If
    
      If lngAmt <> 0 Then
       .TextMatrix(lngRo, cstCol) = Format((lngDiff / lngAmt), cstChangeDisp)
      Else
       .TextMatrix(lngRo, cstCol) = Format(lngAmt, cstChangeDisp)
      End If
    Next lngRo
  End With
  
End Sub
