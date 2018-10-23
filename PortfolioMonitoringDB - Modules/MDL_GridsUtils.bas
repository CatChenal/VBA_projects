Attribute VB_Name = "MDL_GridsUtils"
Option Compare Database
Option Explicit
'================================================================================
'
' MDL GridUtils Nov-20-02 10:45
'
'================================================================================
Const cstMDL = "GridUtils"

Function IsTableInDB(strObjName As String, Optional dbsObj As Variant) As Boolean
  Dim tbl As TableDef
  On Error GoTo IsTableInDBErr
  IsTableInDB = True
  If IsMissing(dbsObj) Then
    CurrentDb.TableDefs.Refresh
    Set tbl = CurrentDb.TableDefs(strObjName)
  Else
    dbsObj.TableDefs.Refresh
    Set tbl = dbsObj.TableDefs(strObjName)
  End If
 
IsTableInDBExit:
  Set tbl = Nothing
  Exit Function
  
IsTableInDBErr:
  If Err <> 3265 Then
    MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation + vbOKOnly, cstMDL & ": IsTableInDB"
  End If
  IsTableInDB = False
  Resume IsTableInDBExit
End Function

Sub ProcessCompFields(dbDAO As DAO.Database, strGridDataTable As String)
' This sub will delete the fields from a stand-alone tbl (strGridDataTable)
' when they are not defined for the company.
'
  Dim tbl As DAO.TableDef
  Dim strCoList$, strFld$
  Dim f As Integer, iFieldsToSkip As Integer
  Dim lngLenFld As Long
  Dim blnDeleted As Boolean
  Dim varPos, varLookup, varCoFields()
  
  f = 0: strCoList$ = "": strFld$ = "": blnDeleted = False: iFieldsToSkip = 2
  On Error GoTo ProcessCompFieldsErr
  
  If lngCurrentComp = 0 Then lngCurrentComp = Forms(cstFRM_Main)!cbxSelComp
  
  ' Get the company defined field names:
  Call GetCoFieldsArray(lngCurrentComp, cstQryCoFlds, varCoFields)
  If IsNull(varCoFields) Then Err.Raise (9)

  ' Process display table according to company-defined fields:
  Set tbl = dbDAO.TableDefs(strGridDataTable)
  If InStr(strGridDataTable, "Form") <> 0 Then iFieldsToSkip = 7
  ' iFieldsToSkip is the number of non-currency fields:
  '   - GridSeries table non-currency fields:
  '       Period End, Months [0, 1]
  '   - SummaryFormData table non-currency fields:
  '       FormCat, Col, ForecastID, ForecastType, DefBudget, PeriodEnd, Months
  '       [   0  ,  1 ,     2     ,       3     ,      4   ,   5      ,    6  ]
  '
  For f = iFieldsToSkip To tbl.Fields.Count - 1
    If f > tbl.Fields.Count - 1 Then Exit For
    strFld$ = tbl.Fields(f).Name
    If Not IsInArray(varCoFields, strFld$, True) Then
      tbl.Fields.Delete strFld$
     'Debug.Print "Deleted fields: " & strFld$
      f = f - 1
    End If
    strFld$ = ""
  Next f
  tbl.Fields.Refresh

ProcessCompFieldsExit:
  Erase varCoFields
  Set tbl = Nothing
  Exit Sub
  
ProcessCompFieldsErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , cstMDL & ": ProcessCompFields"
  Resume ProcessCompFieldsExit
End Sub

Sub CopyFieldValues(lngFrcID As Long, lngCoID As Long, _
                   lngOriginalFldID As Long, lngFldToUpdateID As Long, _
                   blnOriginalFieldIsRequired As Boolean, blnFieldToUpdateIsDefined As Boolean, _
                   blnDeleteOriginalField As Boolean, blnResetToZero As Boolean, _
                   lngResult As Long)
'  lngResult returns:
'                     -1 if an error occurs in any of the in the procs called;
'                      0 if no err;
'                      err.number if err w/in proc.
'
  Dim dbDAO As DAO.Database
  Dim qdf As DAO.QueryDef
  Dim strSQL As String, strOriginalFld As String, strFldToUpdate As String
  Dim lngRecs As Long, lngAcType As Long, lngPrio As Long
  strSQL = "": lngRecs = 0: lngResult = 0
  On Error GoTo CopyFieldValuesErr
         
  Set dbDAO = CurrentDb
  
  If Not blnFieldToUpdateIsDefined Then   'add it to the CompFields tbl
    lngRecs = ExecQuery(dbDAO, "qryAPND_CoField", lngCoID, lngFldToUpdateID)
    If lngRecs < 0 Then 'err in ExecQuery
      lngResult = -1
      Debug.Print "ExecQuery Error: fld " & lngFldToUpdateID & _
                  " possibly not added to Co " & lngCoID
      GoTo CopyFieldValuesExit
    End If
    lngRecs = 0
  End If
  
  ' Get flds table name:
  strOriginalFld = UCase(DLookup("[txtFldTblName]", "tlkpAllFields", "[lngFldId]=" & lngOriginalFldID))
  strFldToUpdate = UCase(DLookup("[txtFldTblName]", "tlkpAllFields", "[lngFldId]=" & lngFldToUpdateID))
  
  ' Create sql string to perform change across series:
  strSQL = "UPDATE tblForecasts INNER JOIN tblSeriesData ON tblForecasts.lngForecastID = "
  strSQL = strSQL & "tblSeriesData.lngForecastID SET tblSeriesData." & strFldToUpdate
  strSQL = strSQL & " = [tblSeriesData]![" & strOriginalFld & "] "
  strSQL = strSQL & "WHERE (tblForecasts.lngForecastID=" & lngFrcID & ");"
  
  ' Create temp action-qry
  Set qdf = dbDAO.CreateQueryDef("")
  qdf.SQL = strSQL
  qdf.Execute dbFailOnError
  lngRecs = qdf.RecordsAffected
  qdf.Close
  Set qdf = Nothing
  dbDAO.TableDefs.Refresh
  
  If blnResetToZero Then  'reset original fld values to 0 in SeriesData tbl
    strSQL = ""
    strSQL = "UPDATE tblForecasts INNER JOIN tblSeriesData ON tblForecasts.lngForecastID ="
    strSQL = strSQL & "tblSeriesData.lngForecastID SET tblSeriesData." & strOriginalFld & "= 0 "
    strSQL = strSQL & "WHERE (tblForecasts.lngForecastID=" & lngFrcID & ");"
    ' Create temp action qry
    Set qdf = dbDAO.CreateQueryDef("")
    qdf.SQL = strSQL
    qdf.Execute dbFailOnError
    qdf.Close
    dbDAO.TableDefs.Refresh
  End If  ' blnResetToZero

  Forms(cstFRM_Fields).Visible = False
  
  If Not blnOriginalFieldIsRequired Then
    If blnDeleteOriginalField Then  'delete from CompFields tbl
      'Close form to avoid record locking error:
      blnCopyField = False
      DoCmd.Close acForm, cstFRM_Fields, acSaveNo
      lngRecs = ExecQuery(dbDAO, "qryDEL_CoField", lngCoID, lngOriginalFldID)
      If lngRecs < 0 Then 'err in ExecQuery
        lngResult = -1
        Debug.Print "ExecQuery Error: fld " & lngFldToUpdateID & _
                    " possibly not deleted from Co " & lngCoID
        GoTo CopyFieldValuesExit
      End If
    End If
  End If  'Not blnOriginalFieldIsRequired
  
  dbDAO.Close
  Call RefreshMainGrids
  
CopyFieldValuesExit:
  Set qdf = Nothing
  Set dbDAO = Nothing
  Exit Sub
  
CopyFieldValuesErr:
  lngResult = Err.Number
  MsgBox "Error (" & Err.Number & "): " & Err.Description, cstMDL & ": CopyFieldValues"
  Resume CopyFieldValuesExit
End Sub

Public Function DiscontinuityExists(lngForecastID As Long, lngCoID As Long, _
                                      strDiscDesc As String) As Boolean
'
' Pre:      If lngForecastID <> 0 then check one forecast only;
'           If lngCoID <>0 (and lngforecastID=0), check each forecast recursively.
'
' Post:     Returned values:
'           Error: True and strDiscDesc holds err.desc;
'           No error:
'            if discontinued: True and strDiscDesc holds a description about
'                             the gap/overlap dates involved;
'                     if not: False and strDiscDesc=OK msg.
'
' Purpose:  To check whether a forecast series has period gap or overlap &
'           notify user if so.
'
' Note:     The overlap date is the most recently entered (see sort order on qry);
'           The gap date is the first missing one.
  Dim dbs As DAO.Database
  Dim qdf As DAO.QueryDef, qdfFrc As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim varData() As Variant
  Dim strSQL As String, strType As String, strDescAll As String
  Dim lngRecs As Long, lngFor As Long
  Dim dteEnd As Date, dtePrevEnd As Date, dteNext As Date
  Dim iMonth As Integer, iMonthDiff As Integer, r As Integer, G As Integer
  Dim blnAll As Boolean, blnProblem As Boolean
  Const cstQryPeriods = "qrySeriesPeriods"      '1 param
  Const cstQryForecasts = "qryCoForecasts"      '2 params
  '-----------------------------------------------------
  blnAll = False:  blnProblem = False: strDiscDesc = ""
  '-----------------------------------------------------
  On Error GoTo DiscontinuityExistsErr
  
  If lngForecastID + lngCoID = 0 Then
    blnProblem = True
    strDiscDesc = "Err: Invalid arguments: the ForecastID and CompID cannot be both zero-valued!"
    Err.Raise 13, "Proc: DiscontinuityExists", strDiscDesc
    GoTo DiscontinuityExistsExit
  End If
  
  Set dbs = CurrentDb
  If lngForecastID <> 0 Then  'use it
  
    Set qdf = dbs.QueryDefs(cstQryPeriods)
    qdf.Parameters(0) = lngForecastID
    Set rst = qdf.OpenRecordset
    If rst.AbsolutePosition = -1 Then
      blnProblem = True
      strDiscDesc = "Err: No Forecasts defined!"
      Err.Raise 17, "Proc: DiscontinuityExists", strDiscDesc
      GoTo DiscontinuityExistsExit
    End If
    rst.MoveLast
    rst.MoveFirst
    lngRecs = rst.RecordCount
    varData = rst.GetRows(lngRecs)
    rst.Close
    Set rst = Nothing
    
    lngRecs = UBound(varData, 2)  'records: 0 to lngRecs
    
    For r = 1 To lngRecs
      'start with the 2nd record in order to ref the previous one in the next Column
      dtePrevEnd = varData(0, r - 1)
      dteEnd = varData(0, r)
      iMonth = varData(1, r)
      strType = UCase(varData(2, r))
      
      iMonthDiff = DateDiff("m", dtePrevEnd, dteEnd)
      
      If iMonthDiff <> iMonth Then  'discontinued
        If blnProblem = False Then   ' on 1st pass print hdr
           strDiscDesc = "Forecast: " & strType & vbCrLf & vbCrLf & _
                        "Type: " & String(9, " ") & "End Date:" & String(5, " ") & _
                         "Periodicity:" & vbCrLf & String(55, "-") & vbCrLf & vbCrLf
        End If
        blnProblem = True
         
        If iMonthDiff > iMonth Then 'GAP
          For G = 1 To (iMonthDiff / iMonth) - 1
            ' Get next month from PrevEnd to list as missing
            dteNext = GetMonthEndDate(DateSerial(Year(dtePrevEnd), Month(dtePrevEnd) + iMonth, 1))
            strDiscDesc = strDiscDesc & "Gap " & String(12, " ") & Format(dteNext, "mm/dd/yy") & _
                          String(10, " ") & iMonth & vbCrLf
            dtePrevEnd = dteNext
          Next G
        Else
          strDiscDesc = strDiscDesc & "Overlap " & String(7, " ") & Format(dteEnd, "mm/dd/yy") & _
                        String(10, " ") & iMonth & vbCrLf
        End If
      End If
    Next r
    
  Else  'use compid
  
    Set qdfFrc = dbs.QueryDefs(cstQryForecasts)
    qdfFrc.Parameters(0) = lngCurrentComp
    qdfFrc.Parameters(1) = Null
    Set rst = qdfFrc.OpenRecordset
    If rst.AbsolutePosition = -1 Then
      strDiscDesc = "Err: No Forecasts defined!"
      Err.Raise 17, "Proc: DiscontinuityExists", strDiscDesc
      GoTo DiscontinuityExistsExit
    End If
    rst.MoveLast
    rst.MoveFirst
    Do While Not rst.EOF
      lngFor = rst(0)
      blnAll = blnAll Or DiscontinuityExists(lngFor, 0, strDiscDesc)
      strDescAll = strDescAll & vbCrLf & strDiscDesc
      rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    qdfFrc.Close
    Set qdfFrc = Nothing
    blnProblem = blnAll
    strDiscDesc = ""
    strDiscDesc = strDescAll
  End If
  
  dbs.Close
  
DiscontinuityExistsExit:
  DiscontinuityExists = blnProblem
  Erase varData
  Set rst = Nothing
  Set qdf = Nothing
  Set qdfFrc = Nothing
  Set dbs = Nothing
  Exit Function
  
DiscontinuityExistsErr:
  blnProblem = True
  strDiscDesc = "Error(" & Err.Number & ") in 'DiscontinuityExists': " & Err.Description
  Resume DiscontinuityExistsExit
End Function

Function CheckDiscontinuity(lngForecastID As Long, lngCoID As Long, _
                            blnShowMessage As Boolean) As Boolean
'
  Dim strDesc As String, str As String, strText As String
  Dim lngErr As Long
  Dim bln As Boolean
  
  strDesc = "": str = ""
  bln = DiscontinuityExists(lngForecastID, lngCoID, strDesc)
  CheckDiscontinuity = bln
  
  ' Check if err
  str = Left(strDesc, 3)
  If str >= "Err" Then
    beep
    strText = "Processing Error:" & vbCrLf & vbCrLf & strDesc
  Else
    If bln Then   'strDesc has description
      strText = "FORECAST DATA POINT DISCONTINUITY:" & vbCrLf & vbCrLf
      strText = strText & strDesc
    Else
      strText = "THIS FORECAST IS CONTINUOUS:" & vbCrLf & vbCrLf
      strText = strText & "There are no date gaps or overlaps."
    End If
  End If
  If blnShowMessage Then
    beep
    DoCmd.OpenForm cstFRM_Msg, , , , acFormReadOnly, , strText
  End If

End Function

Function ExecQuery(dbsDAO As DAO.Database, strQryName As String, _
                                          ParamArray varParams() As Variant) As Long
' Returns RecordsAffected or -1 if error
'
  Dim qdf As DAO.QueryDef
  Dim strErr As String, strTbl As String
  Dim lng As Long
  Dim p As Integer
  lng = 0
  On Error GoTo ExecQueryErr
  
  Set qdf = dbsDAO.QueryDefs(strQryName)
  If qdf.Parameters.Count > 0 Then
    If IsArray(varParams) Then
      For p = 0 To UBound(varParams())
        qdf.Parameters(p).Value = varParams(p)
      Next p
    End If
  End If
  qdf.Execute dbFailOnError
  lng = qdf.RecordsAffected
  qdf.Close
  Set qdf = Nothing
  dbsDAO.TableDefs.Refresh
  
ExecQueryExit:
  ExecQuery = lng
  Set qdf = Nothing
  Exit Function
  
ExecQueryErr:
  If Err = 3010 Then 'tbl to recreate already exists
    strErr = Err.Description
    strTbl = Mid$(Left$(strErr, Len(strErr) - 17), 8)
    'Debug.Print "strErr: " & strErr; "; strTbl: " & strTbl
    Err.Clear
    dbsDAO.TableDefs.Delete strTbl
    dbsDAO.TableDefs.Refresh
    Resume 0
  Else  'If Err > 0 Then
    lng = -1
    MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstMDL & ": ExecQuery"
    Resume ExecQueryExit
  End If
End Function

Public Function Transpose2DArray(varArray As Variant) As Variant
' Custom Function to Transpose a 2-dim, 0-based array

  Dim lngX As Long, lngY As Long, lngXupper As Long, lngYupper As Long
  Dim varTemp As Variant
  On Error GoTo Transpose2DArrayErr
   
  lngYupper = UBound(varArray, 1)   'flds will be in the vertical dir
  lngXupper = UBound(varArray, 2)   'recs -------------- horizontal dir

  ReDim varTemp(lngXupper, lngYupper)
  For lngX = 0 To lngXupper
    For lngY = 0 To lngYupper
      varTemp(lngX, lngY) = varArray(lngY, lngX)
    Next lngY
  Next lngX
  
Transpose2DArrayExit:
  Transpose2DArray = varTemp
  Erase varTemp
  Exit Function
  
Transpose2DArrayErr:
  varTemp = Null
  MsgBox "Error (" & Err.Number & "): " & Err.Description, , cstMDL & ":Transpose2DArray"
  Resume Transpose2DArrayExit
End Function

Public Function IsInArray(varSearchArray As Variant, varSearchValue As Variant, _
                          Optional blnResizeArray As Boolean = False) As Boolean
' Use: Search an array for non-null values (exit on the first match).
' The optional argument permits the 'resizing' of the array so that, when used in a
' loop, it will skip Null values (set on a previous match).
  Dim l As Long
  
  IsInArray = False
  If Not IsArray(varSearchArray) Then Exit Function
  
  If blnResizeArray = True Then
    For l = LBound(varSearchArray) To UBound(varSearchArray)
      If Not IsNull(varSearchArray(l)) Then
        If varSearchValue = varSearchArray(l) Then
          IsInArray = True
          varSearchArray(l) = Null
          Exit For
        End If
      End If
    Next l
  Else
    For l = LBound(varSearchArray) To UBound(varSearchArray)
      If varSearchValue = varSearchArray(l) Then
        IsInArray = True
        Exit For
      End If
    Next l
  End If
End Function

Function CompHasFields(lngCo As Long) As Boolean
  CompHasFields = Not IsNull(DLookup("[lngCompID]", "tblCompFields", "[lngCompID]=" & lngCo))
End Function
