Attribute VB_Name = "CalcMedian"
Option Compare Database
Option Explicit
'
'*****************************************************************
' DF Reports-CalcMedian Sep-24-02 10:35
'******************************************************************
'

 Function GetMedianAmt(intYear As Integer, strAmtFld As String) As Single
'?GetMedianAmt(2000, "O")
  Dim rst As DAO.Recordset
  Dim qdf As DAO.QueryDef
  Dim fldAmt As DAO.Field
  Dim strQry As String
  Dim lngTot As Long
  Dim sglMED As Single
  Dim i As Integer, iOffSet As Integer
  Dim sglOdd As Single, sglVal As Single
        
  If Len(strAmtFld & "") = 0 Then GoTo InvalidName
  
  Select Case Left(strAmtFld, 1)
    Case "O"
      strQry = "qryYearAllAmtOfferedMEDSub"
    Case "T"
      strQry = "qryYearAllTrancheMEDSub"
    Case Else
      GoTo InvalidName
  End Select
  
  Set qdf = CurrentDb.QueryDefs(strQry)
  qdf.Parameters(0) = intYear
  
  Set rst = qdf.OpenRecordset
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
  
GetMedianAmtExit:
  Set fldAmt = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  GetMedianAmt = sglMED
  Exit Function
  
InvalidName:
  MsgBox "GetMedianAmt function not setup for this field. " & _
             "Valid input for field name initial is either 'T' " & _
             "for Tranche Size or 'O' for Amount Offered.", vbCritical, "GetMedianAmt"
  GoTo GetMedianAmtExit
End Function
