Attribute VB_Name = "MDL_ExportUtils"
Option Compare Database
Option Explicit
'================================================================================
'
' MDL_ExportUtils Sep-17-03 12:00
'
'================================================================================
' Names of xl templates:
Public Const cstXLTSummary = "PrintSummaryForm.xlt"
Public Const cstXLTSeries = "PrintSeriesForm.xlt"
  
Const cstMDL = "ExportUtils"
Const cstCancelTitle = "Export Cancelled"

Function CopyGridDataToXLT(frmCallingForm As Form)
  Dim wbk As Excel.Workbook
  '
  Dim dbDAO As DAO.Database
  Dim rst As DAO.Recordset
  Dim ctlGrid As Control
  '
  Dim strWorkbookFullName As String, strCompany As String, strTitle As String 'or forecast desc
  Dim strForm As String, strXLRange As String
  Dim dteSelected As Date
  Dim varZoom As Variant
  '
  Dim lngRows As Long, lngCols As Long
  Dim lngMarginRow As Long, lngMarginRow2 As Long, lngMarginRow3 As Long
  Dim lngCapHrdRow As Long, lngCapHrdRow2 As Long
  
  Dim lngMrktValRow As Long
  Dim lngTotDebtRow As Long, lngTotEqRow As Long, lngTotCapRow As Long, lngFinRatios As Long
  Dim sglMulti As Single
  Dim r As Integer, c As Integer

  ' Names of possible forms:
  Const cstSUM = "sfrmSummary"
  Const cstSER = "sfrmSeries"
  ' Two sheets present in either spreadsheets:
  Const cstForm = "Form" ':what gets printed
  Const cstData = "Data" ':unformatted data.
  Const cstXlRowOffset = 5 ': both xlt have a 5-line header, 1st data row=6th row.
  '-------------------------------------------------------------
  'On Error GoTo CopyGridDataToXLTErr

  strForm = frmCallingForm.Name
     
  ' Open corresponding xl template:
  Select Case strForm
    'Set path and a minimum of variables:
    Case cstSUM
      strWorkbookFullName = GetBackEndDir & cstXLTSummary ' cstCommonXLTPath & cstXLTSummary
      With frmCallingForm
        Set ctlGrid = !ocxFlxGridForm
        dteSelected = !cbxSelPeriodEnd
        sglMulti = !intMultiple
        strCompany = !txtCurrentCompany
        strTitle = !txtSmryTitle
        varZoom = False
      End With
      
    Case cstSER
      strWorkbookFullName = GetBackEndDir & cstXLTSeries
      With frmCallingForm
        .Parent!cbxSelComp.SetFocus
        strCompany = .Parent!cbxSelComp.Column(1)
        strTitle = .Parent!cbxSelForecast.Column(1)
        Set ctlGrid = !ocxFlxGridSeries
        ' Get number of grid rows/cols to iterate:
        lngRows = ctlGrid.Rows
        lngCols = ctlGrid.Cols
        varZoom = 80
      End With
      
    Case Else
      MsgBox "Wrong form name.", vbExclamation, , cstMDL & ": CopyGridDataToXLT"
      GoTo CopyGridDataToXLTExit
  End Select
  Set wbk = GetPFMExcelBook(strWorkbookFullName, True)
  If wbk Is Nothing Then
    MsgBox "Error in GetPFMExcelBook: could not set object", vbExclamation, cstMDL & ": CopyGridDataToXLT"
    GoTo CopyGridDataToXLTExit
  End If
  wbk.RunAutoMacros xlAutoOpen
  wbk.Application.Visible = False
  wbk.Activate

  ' Perform data transfer depending on form:
  Select Case strForm
  
    Case cstSUM         '*************** Export Summary Form data ***************
      Set dbDAO = CurrentDb
      ' Open rst on final tbl:
      Set rst = dbDAO.OpenRecordset(cstFinalTbl)
      rst.MoveLast
      rst.MoveFirst
      ' Save data source row/col count to calculate range:
      lngRows = rst.RecordCount
      lngCols = rst.Fields.Count
      ' Copy unformatted data to the 'data' sheet:
      wbk.Worksheets(cstData).Cells.ClearContents
      strXLRange = "A1:" & wbk.Worksheets(cstData).Cells(lngRows, lngCols).Address
      wbk.Worksheets(cstData).Range(strXLRange).CopyFromRecordset rst
      rst.Close
      '--- End copying from DAO obj.
       
      ' Get number of grid rows/cols to iterate (reusing same var):
      lngRows = ctlGrid.Rows
      lngCols = ctlGrid.Cols
  
      'Copy new input values on Form sheet
      With wbk.Worksheets(cstForm) 'reset wsh object to 'Form' sheet
     
        .Activate
        .Range("CoName") = strCompany
        .Range("EndDate") = dteSelected
        .Range("SummaryTitle") = strTitle
        .Calculate
        
        For r = 1 To lngRows - 1
          For c = 1 To lngCols
            If c = 1 Then
              Select Case ctlGrid.TextMatrix(r, c - 1)
                Case "Margins"
                  lngMarginRow = r + cstXlRowOffset
                Case "Tot Debt"
                  lngTotDebtRow = r + cstXlRowOffset
                Case "Tot Equity"
                  lngTotEqRow = r + cstXlRowOffset
                Case "Tot Capitalization"
                  lngTotCapRow = r + cstXlRowOffset
              End Select
            ElseIf c = 11 Then
              If Left(ctlGrid.TextMatrix(r, c - 1), 3) = "Est" Then lngMrktValRow = r + cstXlRowOffset
              If Left(ctlGrid.TextMatrix(r, c - 1), 3) = "FIN" Then lngFinRatios = r + cstXlRowOffset
            End If
            .Cells(r + cstXlRowOffset, c) = ctlGrid.TextMatrix(r, c - 1)
          Next c
        Next r
        
        'Format
        strXLRange = "A" & lngMarginRow   '1st mrg row
        lngMarginRow2 = lngMarginRow + 2 '2 others
        strXLRange = strXLRange & ":A" & lngMarginRow2
        .Range(strXLRange).HorizontalAlignment = xlRight
        .Range(strXLRange).Font.Bold = True
        
        lngCapHrdRow = lngMarginRow2 + 2 '1st cap hdr row
        strXLRange = "A" & lngCapHrdRow
        .Range(strXLRange).HorizontalAlignment = xlRight
        .Range(strXLRange).Font.Bold = True
       
        lngCapHrdRow2 = lngCapHrdRow + 1 '2nd cap hdr row
        .Range("H" & lngCapHrdRow & ":I" & lngCapHrdRow2).Font.Underline = True
        .Range("M" & lngCapHrdRow & ":N" & lngCapHrdRow2).Font.Underline = True
        .Range(strXLRange).HorizontalAlignment = xlLeft
        strXLRange = strXLRange & ":P" & lngCapHrdRow2
        .Range(strXLRange).Font.Bold = True
    
        strXLRange = "A" & lngTotDebtRow
        .Range(strXLRange).HorizontalAlignment = xlRight
        .Range(strXLRange).Font.Bold = True
       
        strXLRange = "A" & lngTotEqRow
        .Range(strXLRange).HorizontalAlignment = xlRight
        .Range(strXLRange).Font.Bold = True
         
        strXLRange = "A" & lngTotCapRow
        .Range(strXLRange).HorizontalAlignment = xlRight
        .Range(strXLRange).Font.Bold = True
         
        If lngMrktValRow > 0 Then
          strXLRange = "L" & lngMrktValRow
          .Range(strXLRange).HorizontalAlignment = xlRight
        End If
         
        strXLRange = "L" & lngFinRatios
        .Range(strXLRange).Font.Bold = True
        
        'Resize largest columns:
        .Columns("G").ColumnWidth = 18
        .Columns("K").ColumnWidth = 18
        
        ' Reset print range
        strXLRange = "$A$1:$P$" & lngTotCapRow + 1   'lngRows
      End With
     
    Case cstSER  '*************** Export Series Form data ***************
      ' Copy unformatted data to the 'data' sheet:
      With wbk.Worksheets(cstData)
        'Copy grid data:
        For r = 0 To lngRows - 1
          For c = 0 To lngCols - 1
            .Cells(r + 1, c + 1) = ctlGrid.TextMatrix(r, c)
          Next c
        Next r
        'Get range of data area:
        strXLRange = "A1:" & .Cells(lngRows, lngCols).Address
        .Range(strXLRange).Copy
        ' Redefine pasting range:
        strXLRange = .Range(strXLRange).Offset(rowOffset:=cstXlRowOffset).Address
        .Paste Destination:=wbk.Worksheets(cstForm).Range(strXLRange)
      End With
      '
      With wbk.Worksheets(cstForm) 'reset wsh object to 'Form' sheet
        .Activate
        .Range("CoName") = strCompany
        .Range("FRC") = strTitle
        .Calculate
        ' Reset print range:
        strXLRange = "A1:" & Mid$(strXLRange, 6) 'include header
      End With
  End Select
  '
  wbk.Application.Visible = True  'if not wbk stays hidden
  
  With wbk.Worksheets(cstForm)
    .Range(strXLRange).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    .Columns(1).Font.Bold = True
    .Rows(6).Font.Bold = True
    .Rows(6).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    .PageSetup.PrintArea = cstForm & "!" & strXLRange
    .PageSetup.Zoom = varZoom
  End With
  wbk.PrintPreview True
  
CopyGridDataToXLTExit:
  Set rst = Nothing
  Set dbDAO = Nothing
  Set wbk = Nothing
  Exit Function
  
CopyGridDataToXLTErr:
  Screen.MousePointer = 0
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , cstMDL & ": CopyGridDataToXLT"
  Resume CopyGridDataToXLTExit
End Function

Function ExportForecast(lngFrcID As Long, lngCoID As Long, lngFrcTypeID As Long, _
                        wbkXL As Excel.Workbook)
' Called from the Export Selection form
' The forecasts selected for export will correspond to the non-zero argument.
' i.e.: to export one forecast, use lngFRCID (<>0);
'       to export all of a company's forecasts, use lngCoID (<>0)
'       to export specific types use lngFrcTypeID:
'           0=all; 1=Actuals, 2=Budgets;3=Downside;4=Forecast;5=Basecase.
'
  Dim wsh As Excel.Worksheet
  Dim rstFo As DAO.Recordset, rstCo As DAO.Recordset
  Dim qdfFo As DAO.QueryDef, qdfCo As DAO.QueryDef
  Dim strMsg As String, strWorkbookFullName As String, strFRCDesc As String ': will be the sheet name
  Dim strRangeOutput As String
  Dim lngFor As Long
  Dim lngRows As Long, lngCols As Long, c As Long
  
  Const cstQryCoFRC = "qryCoForecasts"
  Const cstQryFRCData = "qryForecastData"
  Const cstExportTbl = "tblExportFRCST"
  '------------------------------------------------------------------
  lngFor = 0: strMsg = ""
  '------------------------------------------------------------------
  On Error GoTo ExportForecastErr
  
  ' Check args:
  If lngFrcID + lngCoID = 0 Then 'no go
    strMsg = "Invalid arguments: the ForecastID and CompID cannot be both zero-valued!"
    GoTo ExportForecastExit
  End If
  If lngFrcID * lngCoID > 0 Then 'no go
    strMsg = "Invalid arguments: either one of the ForecastID and CompID need to be set to 0!"
    GoTo ExportForecastExit
  End If
  
  If lngFrcID > 0 Then 'use it
  
    Set qdfFo = CurrentDb.QueryDefs(cstQryFRCData)
    qdfFo.Parameters(0) = lngFrcID
    qdfFo.Parameters(1) = lngFrcTypeID
    Set rstFo = qdfFo.OpenRecordset
    If rstFo.AbsolutePosition = -1 Then
      strMsg = "There is no data point associated with this forecast."
      rstFo.Close
      qdfFo.Close
      GoTo ExportForecastExit
    End If
    rstFo.MoveLast
    rstFo.MoveFirst
    Call CreateTransposedFRCTable(rstFo, cstExportTbl, strFRCDesc)
    rstFo.Close
    Set rstFo = Nothing
    qdfFo.Close
    Set qdfFo = Nothing
    
    ' Reopen against new tbl:
    Set rstFo = CurrentDb.OpenRecordset(cstExportTbl)
    rstFo.MoveLast
    rstFo.MoveFirst
    lngRows = rstFo.RecordCount
    lngCols = rstFo.Fields.Count
    
    Set wsh = wbkXL.Worksheets.Add    'sheet that will receive the transposed rst data
    On Error Resume Next
    wsh.Name = strFRCDesc
    If Err.Number <> 0 Then
      If Err.Number = 1004 Then 'name not unique
        Err.Clear
        strFRCDesc = strFRCDesc & "-" & wbkXL.Worksheets.Count
        wsh.Name = strFRCDesc
        Resume Next
      Else
        Debug.Print "ExportForecast Err.Number: " & Err.Number
        Resume ExportForecastErr
      End If
    End If
    
    wsh.Activate
    With wsh
      For c = 0 To lngCols - 1
        .Cells(1, c + 1).Value = rstFo.Fields(c).Name
      Next
      .Range(.Cells(1, 1), .Cells(1, lngCols)).Font.Bold = True

      strRangeOutput = "A2:" & .Cells(lngRows, lngCols).Address
      .Range(strRangeOutput).CopyFromRecordset rstFo
    End With

    rstFo.Close

  Else
  
    Set qdfCo = CurrentDb.QueryDefs(cstQryCoFRC)
    qdfCo.Parameters(0) = lngCoID
    qdfCo.Parameters(1) = lngFrcTypeID
    Set rstCo = qdfCo.OpenRecordset
    If rstCo.AbsolutePosition = -1 Then
      strMsg = cstNoFrcForThisComp
      rstCo.Close
      qdfCo.Close
      GoTo ExportForecastExit
    End If
    rstCo.MoveLast
    rstCo.MoveFirst
    Do While Not rstCo.EOF
      lngFor = rstCo(0)
      Call ExportForecast(lngFor, 0, lngFrcTypeID, wbkXL)
      rstCo.MoveNext
    Loop
    rstCo.Close
    qdfCo.Close
  End If
  
ExportForecastExit:
  If Len(strMsg) > 0 Then
    MsgBox strMsg, vbExclamation, cstCancelTitle
  End If
  DoCmd.Hourglass False
  Set qdfFo = Nothing
  Set qdfCo = Nothing
  Set rstFo = Nothing
  Set rstCo = Nothing
  Set wsh = Nothing
  Exit Function
  
ExportForecastErr:
  strMsg = ""
  'wbkXL.Application.Cursor = xlDefault
  Set wbkXL = Nothing
  Debug.Print "ExportForecast Err.Number: " & Err.Number
  MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, cstMDL & ": ExportForecast"
  Resume ExportForecastExit

End Function

Function GetPFMExcelBook(strBookName As String, blnShowApp As Boolean) As Excel.Workbook
' Proc opens a new workbook if strBookName=""
'
  Dim appXL As Excel.Application
  Dim xlBook As Excel.Workbook
  Dim blnAppAlreadyRunning  As Boolean
  DoCmd.Hourglass True
  '------------------------------------------------------
  On Error Resume Next
  
  Set appXL = AppOpen("XLMain", "Excel.Application", True, blnAppAlreadyRunning)
  blnExcelAlreadyRunning = blnExcelAlreadyRunning Or blnAppAlreadyRunning
  With appXL
    If Len(strBookName) = 0 Then
      Set xlBook = .Workbooks.Add
    Else
      Set xlBook = .Workbooks.Open(strBookName, , , , , , , , , , , , False)
    End If
    If Err <> 0 Then
      Set xlBook = Nothing
      If Err = 1004 Then
        MsgBox "The template was not found in this directory (where it should be): " & _
           vbCrLf & GetBackEndDir, vbExclamation, "GetPFMExcelBook"
           GoTo GetPFMExcelBookExit
      ElseIf Err = 91 Then
         MsgBox "Cannot set xlBook variable ", vbExclamation, "GetPFMExcelBook"
          GoTo GetPFMExcelBookExit
      Else
        GoTo GetPFMExcelBookErr
      End If
    End If
    .Visible = blnShowApp
    If blnShowApp Then .WindowState = xlNormal
  End With
  
GetPFMExcelBookExit:
  Set GetPFMExcelBook = xlBook
  DoCmd.Hourglass False
  Set appXL = Nothing
  Exit Function
  
GetPFMExcelBookErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, "GetPFMExcelBook"
  Resume GetPFMExcelBookExit
End Function

Sub ExportForecastPrep(frm As Form)
  Dim wbk As Excel.Workbook
  Dim qdfAll As DAO.QueryDef
  Dim rstAll As DAO.Recordset
  Dim intResponse As Integer
  Dim strBkName As String, strMsg As String
  Dim lngFoID As Long, lngCoID As Long, lngFRC As Long
  Dim blnAll As Boolean
  blnAll = False
  
LocateFile:
  intResponse = 0
  intResponse = MsgBox("Do you want to export to a new workbook?", vbQuestion + vbYesNoCancel, _
                       "Export Destination")
  If intResponse = vbCancel Then
     GoTo ExportForecastPrepExit
  ElseIf intResponse = vbNo Then
    'Call API function to show Open File window and return name selected
    strBkName = ReturnOpenFileName(frm, "*.xls" & Chr(0) & "*.XLS")
  ElseIf intResponse = vbYes Then
    strBkName = ""
  End If
  
  If Len(strBkName) = 0 And intResponse = vbNo Then
    intResponse = MsgBox("You have not specified a destination file!", vbExclamation + vbRetryCancel)
    If intResponse = vbCancel Then
      GoTo ExportForecastPrepExit
    Else
      GoTo LocateFile
    End If
  End If
  
  lngFoID = lngCurrentForecast 'frm!lngFrcID
  lngCoID = lngCurrentComp  'frm!lngCoID
  lngFRC = 0
  If frm!cbxSelFrcType.Enabled Then lngFRC = frm!cbxSelFrcType.Column(0)

  Select Case frm!opgExportSel
    ' Reset unneeded parameter(s) before passing them to export fct
    Case 0  'all forecasts
      lngFoID = 0: lngCoID = 0: blnAll = True
    Case 1  'all of current comp forecasts
      lngFoID = 0
    Case 2  'current forecast only
      lngCoID = 0
    Case Else
      strMsg = "Export operation cancelled by user."
      GoTo ExportForecastPrepExit
  End Select
  
  frm.Visible = False
  Set wbk = GetPFMExcelBook(strBkName, True) 'if strBookName = "": open a new book
  If wbk Is Nothing Then
    strMsg = "Error in GetPFMExcelBook: could not set object"
    GoTo ExportForecastPrepExit
  End If
  DoCmd.Hourglass True
  wbk.Application.Cursor = xlWait
  
  If blnAll Then
    Set qdfAll = CurrentDb.QueryDefs("qryAllForecasts")
    qdfAll.Parameters(0) = lngFRC
    Set rstAll = qdfAll.OpenRecordset
    rstAll.MoveLast
    rstAll.MoveFirst
    With rstAll
      Do While Not .EOF
        lngFoID = !lngForecastID
        ' Call to ExportForecast function:
        Call ExportForecast(lngFoID, 0, lngFRC, wbk)
       .MoveNext
      Loop
      .Close
    End With
    qdfAll.Close
    Set qdfAll = Nothing
    Set rstAll = Nothing
  Else
    ' Call to ExportForecast function
    Call ExportForecast(lngFoID, lngCoID, lngFRC, wbk)
  End If
  '
  If Not (wbk Is Nothing) Then
    wbk.Application.Cursor = xlDefault
    wbk.Application.WindowState = xlNormal
  End If
  frm.Visible = True
  beep
  MsgBox "No more forecasts to export.", vbInformation, "Done"
    
ExportForecastPrepExit:
  If Len(strMsg) > 0 Then
    MsgBox strMsg, vbExclamation, cstCancelTitle
  End If
  DoCmd.Hourglass False
  Set qdfAll = Nothing
  Set rstAll = Nothing
  Set wbk = Nothing
  DoCmd.Close acForm, frm.Name
  Set frm = Nothing
  Exit Sub
  
ExportForecastPrepErr:
  If Not (frm Is Nothing) Then frm.Visible = True
  MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, cstMDL & ": ExportForecastPrep"
  Resume ExportForecastPrepExit
End Sub

Function CreateTransposedFRCTable(rstSource As DAO.Recordset, strTarget As String, _
                                  strFRCDesc As String)
 ' Creats a transposed tbl to be used to export a selected forecast to xl.
  Dim tdfNew As DAO.TableDef
  Dim fldNew As DAO.Field
  Dim rstTarget As DAO.Recordset
  Dim strFldName As String
  Dim i As Integer, j As Integer
  Dim sglVal As Single
  '---------------------------------------------------------------
  i = 0: j = 0:  sglVal = 0
  ' Note: rstSource is from cstQryFRCData = "qryForecastData":
  ' First field: frc desc (dbText) will provide the sheet name.
  ' Second     : date field to become fld name (col)
  ' All others : dbsingle
  '---------------------------------------------------------------
  On Error GoTo CreateTransposedFRCTableErr
 
  rstSource.MoveLast
  rstSource.MoveFirst
  strFRCDesc = ""
  strFRCDesc = rstSource(0) 'save description
  ' Create a new table to hold the transposed data.
  ' Create a field for each record in the original recordset.
  If IsTableInDB(strTarget) Then CurrentDb.TableDefs.Delete strTarget

  Set tdfNew = CurrentDb.CreateTableDef(strTarget)
  
  ' First fld is text
  Set fldNew = tdfNew.CreateField("Desc", dbText)
  tdfNew.Fields.Append fldNew
  
  For i = 0 To rstSource.RecordCount - 1
    strFldName = CStr(rstSource.Fields(3).Value)
    Set fldNew = tdfNew.CreateField(strFldName, dbSingle)
    tdfNew.Fields.Append fldNew
    rstSource.MoveNext
  Next i
  Set fldNew = Nothing
   
  ' Append the table to the TableDefs.
  CurrentDb.TableDefs.Append tdfNew
  Set tdfNew = Nothing
  
  rstSource.MoveFirst
  ' Open the new table and fill the first field with field names from the original table.
  Set rstTarget = CurrentDb.OpenRecordset(strTarget)
  ' Fill each column of the new table with a record from the original table.
  ' Skip first 2 flds: SeriesID, ForecastID
  For j = 4 To rstSource.Fields.Count - 1
    rstTarget.AddNew
    
    For i = 0 To rstTarget.Fields.Count - 1
    
      With rstTarget
        If i = 0 Then
          .Fields(i) = rstSource.Fields(j).Properties("Caption")
        Else
          If j > 0 Then
            If Len(rstSource.Fields(j) & "") = 0 Then
              sglVal = 0
            Else
              sglVal = CLng(rstSource.Fields(j))
            End If
            .Fields(i) = sglVal
          End If
          rstSource.MoveNext
        End If
      End With
      
    Next i
    rstTarget.Update
    rstSource.MoveFirst
  Next j
  
  rstTarget.Close
  Set rstTarget = Nothing
  CurrentDb.TableDefs.Refresh

CreateTransposedFRCTableExit:
  Set rstTarget = Nothing
  Exit Function
  
CreateTransposedFRCTableErr:
  Select Case Err
    Case 3010
       MsgBox "The table " & strTarget & " already exists."
    Case 3078
       MsgBox "The recordset doesn't exist."
    Case Else
       MsgBox Err & " " & Err.Description, vbExclamation, cstMDL & ": CreateTransposedFRCTable"
  End Select
  Resume CreateTransposedFRCTableExit
End Function
