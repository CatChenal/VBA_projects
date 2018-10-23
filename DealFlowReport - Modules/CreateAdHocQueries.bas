Attribute VB_Name = "CreateAdHocQueries"
Option Compare Database
Option Explicit
'
'*****************************************************************
' DF Reports-CreateAdHocQueries
' Updated: Oct-01-03 16:55
'******************************************************************
'

Function SQL_1_CountOfDealsWith1Sec(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "DealData"    'Xqry2001DealData
  'cstNoWarrantsQ  = " & cstNoWarrantsQ & "

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".lngDealNum,  " & cstNoWarrantsQ & ".lngSecDealNum, " & _
  "Count( " & cstNoWarrantsQ & ".lngSecNum) AS [Sec Count] FROM " & strSubQry1 & " INNER JOIN  " & _
  cstNoWarrantsQ & " ON " & strSubQry1 & ".lngDealNum =  " & cstNoWarrantsQ & ".lngSecDealNum " & _
  "GROUP BY " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".lngDealNum,  " & cstNoWarrantsQ & ".lngSecDealNum " & _
  "HAVING (((Count( " & cstNoWarrantsQ & ".lngSecNum))<=1));"

  strQry = "qry" & lngYr & "CountOfDealsWith1Sec"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  
End Function

Function SQL_1_DealsWith0Sec(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "DealData"           'xqry2001DealData

  strSQL = "SELECT  " & strSubQry1 & ".SelectedYear,  " & strSubQry1 & ".lngDealNum, Count( " & _
    cstNoWarrantsQ & ".lngSecDealNum) AS [Sec Count] FROM  " & strSubQry1 & " LEFT JOIN " & _
    cstNoWarrantsQ & " ON  " & strSubQry1 & ".lngDealNum =  " & cstNoWarrantsQ & ".lngSecDealNum " & _
    "GROUP BY  " & strSubQry1 & ".SelectedYear,  " & strSubQry1 & ".lngDealNum " & _
    "HAVING (((Count( " & cstNoWarrantsQ & ".lngSecDealNum))<1));"
  
  strQry = "qry" & lngYr & "DealsWith0Sec"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQL_1_DealsWith1Sec(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "DealData"           'xqry2001DealData
  'cstNoWarrantsQ  =  " & cstNoWarrantsQ & "

  strSQL = "SELECT  " & strSubQry1 & ".SelectedYear,  " & strSubQry1 & ".lngDealNum,  " & strSubQry1 & ".lngDealIssuerNum, " & _
      "Count( " & cstNoWarrantsQ & ".lngSecDealNum) AS [Sec Count] " & _
      "FROM  " & strSubQry1 & " INNER JOIN  " & cstNoWarrantsQ & " ON  " & strSubQry1 & ".lngDealNum = " & _
      cstNoWarrantsQ & ".lngSecDealNum GROUP BY  " & strSubQry1 & ".SelectedYear, " & _
      strSubQry1 & ".lngDealNum,  " & strSubQry1 & ".lngDealIssuerNum " & _
      "HAVING (((Count( " & cstNoWarrantsQ & ".lngSecDealNum))=1));"
  
  strQry = "qry" & lngYr & "DealsWith1Sec"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQL_1_DealsWith4PlusSec(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "DealData"           'xqry2001DealData
  'cstNoWarrantsQ  =  " & cstNoWarrantsQ & "

  strSQL = "SELECT   " & strSubQry1 & ".SelectedYear,   " & strSubQry1 & ".lngDealNum,   " & strSubQry1 & ".lngDealIssuerNum, " & _
  "Count( " & cstNoWarrantsQ & ".lngSecDealNum) AS [Sec Count] FROM   " & strSubQry1 & " INNER JOIN " & _
  cstNoWarrantsQ & " ON   " & strSubQry1 & ".lngDealNum =  " & cstNoWarrantsQ & ".lngSecDealNum " & _
  "GROUP BY   " & strSubQry1 & ".SelectedYear,   " & strSubQry1 & ".lngDealNum,   " & strSubQry1 & ".lngDealIssuerNum " & _
  "HAVING (((Count( " & cstNoWarrantsQ & ".lngSecDealNum))>=4));"
  
  strQry = "qry" & lngYr & "DealsWith4PlusSec"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function


