Attribute VB_Name = "MDL CalcMedian"
Option Compare Database
Option Explicit
'================================================================================
' DBFrontEnd
' MDL CalcMedian Aug-29-02
'
'================================================================================
Const cstMDL = "CalcMedian"

Function GetMedianAmt(intYear As Integer, strAmtFld As String, strFormat As String) As Variant
  '?GetMedianAmt(2000, "O", "$#,###.000")
  Dim rst As DAO.Recordset
  Dim qdf As DAO.QueryDef
  Dim fldAmt As DAO.Field
  Dim param As DAO.Parameter
  Dim frmFilt As Form
  Dim strQry As String, strParam As String, strCtl As String
  Dim lngTot As Long, lngLastExclamPt As Long
  Dim sglMED As Single, sglOdd As Single, sglVal As Single
  Dim blnFormat As Boolean
  Dim i As Integer, iOffSet As Integer, p As Integer, iParams As Integer

  sglMED = 0
  On Error GoTo GetMedianAmtErr
    
  If Not IsLoaded(cstFilterForm) Then GoTo GetMedianAmtExit
  Set frmFilt = Forms(cstFilterForm)

  If Len(strAmtFld & "") = 0 Then GoTo InvalidName
  i = InStr("OTI", Left(strAmtFld, 1))
  Select Case i
    Case 1  'offered
      strQry = "qryYearAllAmtOfferedMEDSub"
    Case 2 'tranche
      strQry = "qryYearAllTrancheMEDSub"
    Case 3 'invested
      strQry = "qryYearAllAmtInvestedMEDSub"
    Case Else
      GoTo InvalidName
  End Select
  i = 0 'reset
  
  blnFormat = Len(strFormat & "") <> 0
    
  Set dbs = CurrentDb
  Set qdf = dbs.QueryDefs(strQry)
  iParams = qdf.Parameters.Count - 1
  For p = 0 To iParams
    Set param = qdf.Parameters(p)
    ' Processing to avoid fixed reference of param indeces
    If p <> iParams Then  'not last param: originating from subquery
      ' Format of param name=[Forms]![frmFilterForm]![cbxSelAnalyst]
      strParam = param.Name
      lngLastExclamPt = InStr(10, strParam, "!") + 1 ' start past the first eclamation pt
      strCtl = Mid$(strParam, lngLastExclamPt)
      ' Reference the corresponding frmFilt field for its value:
      param.Value = frmFilt.Controls(strCtl)
    Else
      param.Value = intYear
    End If
  Next p
  
  Set rst = qdf.OpenRecordset()
  If rst.AbsolutePosition <> -1 Then
    Set fldAmt = rst.Fields(1)  'second field defined in qry
    With rst
      .MoveLast
  
      lngTot = .RecordCount
      sglOdd = lngTot Mod 2
      
      If sglOdd <> 0 Then
        iOffSet = ((lngTot + 1) / 2) - 2
        For i = 0 To iOffSet
          .MovePrevious
        Next i
        sglMED = fldAmt.Value
      Else
        iOffSet = (lngTot / 2) - 2
        For i = 0 To iOffSet
          .MovePrevious
        Next i
        sglOdd = fldAmt.Value
        .MovePrevious
        sglVal = fldAmt.Value
        sglMED = (sglOdd + sglVal) / 2
      End If
      .Close
    End With
  End If
  qdf.Close
  dbs.Close
  
GetMedianAmtExit:
  Set param = Nothing
  Set frmFilt = Nothing
  Set fldAmt = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  
  GetMedianAmt = IIf(blnFormat, Format(sglMED, strFormat), sglMED)
  Exit Function
  
InvalidName:
  MsgBox "GetMedianAmt function not setup for this field. " & vbCrLf & _
         "The median is calculated for the Tranche, Offered and Invested amounts." & _
         vbExclamation, "GetMedianAmt"
  GoTo GetMedianAmtExit
  
GetMedianAmtErr:
  sglMED = 0
  MsgBox "Error: (" & Err & ") " & Err.Description, vbExclamation, cstMDL & " GetMedianAmt"
  Resume GetMedianAmtExit
End Function
