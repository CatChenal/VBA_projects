Attribute VB_Name = "CreateReportQueries"
Option Compare Database
Option Explicit
'
'*****************************************************************
' DF Reports-CreateReportQueries
' Update: Mar-29-04 Updated LTM-Qx queries.
' Update: Jan-09-04 Updated yearly queries.
' Update: Jan-06-04 Added YearClosed to SQLQr_LTMQEndClosedDealsPerMonth
'                    Changed relationship in SQLQr_LTMQEndClosedVsReviewedDeals (SelectedYear==YearClosed)
'                    Added YearClosed to SQLYr_TotClosedDealsForSingleDealSources
'                    Added SelectedQ to SQLYr_IssuersFinData
'                     Removed SelectedQ from SQLYr_ShareOfCoverageLevel
'                     Removed SelectedQ from SQLYr_ShareOfLeverageLevel
'
' Update: Oct-30-03 10:10
'          Corrected error in LTM-Q3 range
'******************************************************************
'
Public Const cstNoWarrantsQ = "qryNonWarrantSecuritiesData"
Public Const cstQryDescY = "Yearly charts & tables query"
Public Const cstQryDescQ = "Quarterly charts & tables query"
Public Const cstDescAll = "Shared query - all years"
'''
Const cstClosedDealsSummAll = "tblClosedDealsSummary-All"
Const cstClosedDealsSummAllMTBL = "qryClosedDealsSummaryMTBL"
Const cstSourcesWithDeals = "qrySourcesWithDealID"
'''
Const cstClosedDealsSumm = "qryClosedDealsSummary"
Const cstClosedDealsSummMTBL = "qryClosedDealsSummaryMTBL"
'''
Public blnOverwrite As Boolean

Function CreateQuarterlyQueries(lngYr As Long, iQ As Integer)
'?CreateQuarterlyQueries(2002, 1)
  If Not blnOverwrite Then  'ask
    blnOverwrite = (vbYes Mod (MsgBox("Do you want to overwrite existing Q" & iQ & " queries?", _
                                    vbQuestion + vbYesNo, "Overwrite Quarterly Queries?")) = 0)
  End If
  
  Call SQLQr_DealData(lngYr, iQ)
  Call SQLQr_TotReviewedDeals(lngYr, iQ)
  Call SQLQr_AggregInvAndAverSize(lngYr, iQ)
  
  If lngYr > 1998 Then
    Call CreateLTMQxQueries(lngYr, iQ)
    Call SQLQr_UNPrevYearQEndDealSizeComp(lngYr, iQ)
  End If
  '---
  Call SQLQr_ClosedDealData(lngYr, iQ)
  Call SQLQr_TotClosedDeals(lngYr, iQ)
  Call RunDBQuery(cstClosedDealsSummMTBL)
  '---
  Call SQLQr_IssuersFinData(lngYr, iQ)
  Call SQLQr_CountOfIssuers(lngYr, iQ)
  Call SQLQr_IssuersWithFinData(lngYr, iQ)
  Call SQLQr_CountOfIssuersWithFinData(lngYr, iQ)
  '---
  Call SQLQr_IssuerCoverage(lngYr, iQ)
  Call SQLQr_ShareOfCoverageLevel(lngYr, iQ)
  Call SQLQr_IssuerLeverage(lngYr, iQ)
  Call SQLQr_ShareOfLeverageLevel(lngYr, iQ)
  '---
  Call SQLQr_CountOfDealType(lngYr, iQ)
  Call SQLQr_AverSizePerDealType(lngYr, iQ)
  Call SQLQr_ShareOfDealType(lngYr, iQ)
  '---
  Call SQLQr_CountOfSecType(lngYr, iQ)
  Call SQLQr_AverSizePerSecType(lngYr, iQ)
  Call SQLQr_ShareOfSecType(lngYr, iQ)
  '---
  Call SQLQr_CountOfSourceType(lngYr, iQ)
  Call SQLQr_ShareOfSourceType(lngYr, iQ)
  '---
  Call SQLQr_SourcesWith1DealSub(lngYr, iQ)
  Call SQLQr_SourcesWith1Deal(lngYr, iQ)
  Call SQLQr_TotSourcesWith1Deal(lngYr, iQ)
  Call SQLQr_TotClosedDealsForSingleDealSources(lngYr, iQ)
  Call SQLQr_SourcesWith1DealAverSize(lngYr, iQ)
  Call SQLQr_SummarySingleDealSources(lngYr, iQ)
  '---
  
  Call SQLQr_SourcesWith2PlusDeals(lngYr, iQ)
  Call SQLQr_TotSourcesWith2PlusDeals(lngYr, iQ)
  Call SQLQr_TotClosedDealsForMultiDealSources(lngYr, iQ)
  Call SQLQr_SourcesWith2PlusDealsSecData(lngYr, iQ)
  Call SQLQr_SummaryMultiDealSources(lngYr, iQ)
  Call SQLQr_UNSourcesSummary(lngYr, iQ)
  
  Application.CurrentDb.QueryDefs.Refresh
  
  MsgBox "All quarterly queries created for Q" & iQ & "-" & lngYr, , "CreateQuarterlyQueries"

End Function

Sub SQLQr_TotReviewedDeals(lngYr As Long, iQ As Integer)
  Dim strQry As String, strSQL As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQ & "DealData"
           
  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & _
        ".SelectedQ, Count(" & strSubQry1 & ".lngDealNum) AS ReviewedDeals FROM " & strSubQry1 & _
        " GROUP BY " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ" & _
        " ORDER BY Count(" & strSubQry1 & ".lngDealNum);"
  
  strQry = "qry" & lngYr & "Q" & iQ & "TotReviewedDeals"
  Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
    
End Sub

Sub CreateGlobalQueries()
  blnOverwrite = (vbYes Mod (MsgBox("Do you want to overwrite existing queries?", _
                                    vbQuestion + vbYesNo, "Overwrite Global Queries?")) = 0)

  ' All years queries
  Call SQL_ClosedDealData
  Call SQL_IssuersFinData
  Call SQL_IssuersWithFinData
  Call SQL_InvestmentAmtNonAddOnDeals
  Call SQL_SourcesWithDealID
  Call SQL_ClosedDealsSummaryAllMTBL
  Call SQL_ClosedDealsSummary
  Debug.Print "CreateGlobalQueries: All Years queries created."
End Sub

Function CreateUNSourcesSummary(lngYr As Long)
  blnOverwrite = (vbYes Mod (MsgBox("Do you want to overwrite existing queries?", _
                                    vbQuestion + vbYesNo, "Overwrite Source Summary Queries?")) = 0)

  Call SQLYr_SourcesWith1DealSub(lngYr)
  Call SQLYr_SourcesWith1Deal(lngYr)
  Call SQLYr_TotSourcesWith1DealSub(lngYr)
  Call SQLYr_TotSourcesWith1Deal(lngYr)
  Call SQLYr_TotClosedDealsForSingleDealSources(lngYr)
  Call SQLYr_SourcesWith1DealAverSize(lngYr)
  Call SQLYr_SummarySingleDealSources(lngYr)
  '---
  Call SQLYr_SourcesWith2PlusDeals(lngYr)
  Call SQLYr_SourcesWith2PlusDealsSecData(lngYr)
  Call SQLYr_TotClosedDealsForMultiDealSources(lngYr)
  Call SQLYr_SummaryMultiDealSources(lngYr)
  '---
  Call SQLYr_UNSourcesSummary(lngYr)
End Function

Function CreateYearlyQueries(lngYr As Long, Optional blnOverwriteQry As Boolean = True)
'?CreateYearlyQueries(2002)
  Dim Q As Integer
  
  blnOverwrite = blnOverwriteQry
  
  Call SQLYr_DealData(lngYr)
  Call SQLYr_ClosedDealData(lngYr)
  Call SQLYr_NonWarrantSecuritiesData
  Call SQLYr_SourcesWithDealID(lngYr)
  
  Call SQLYr_IssuersFinData(lngYr)
  Call SQLYr_IssuersWithFinData(lngYr)
  Call SQLYr_CountOfIssuersWithFinData(lngYr)
  Call SQLYr_ClosedDealsSummaryMTBL(lngYr)    'execute
  Call SQLYr_ClosedDealsSummary(lngYr)
  
  ' Create next level of dependent qries
  Call SQLYr_IssuerCoverage(lngYr)
  Call SQLYr_IssuerLeverage(lngYr)
'
  Call SQLYr_ClosedDealsPerMonth(lngYr)
  Call SQLYr_ReviewedDealsPerMonth(lngYr)
  Call SQLYr_TotReviewedDeals(lngYr)
  Call SQLYr_TotClosedDeals(lngYr)
  Call SQLYr_ClosedDealsSummary(lngYr)
  '---
  Call SQLYr_CountOfDealType(lngYr)
  Call SQLYr_CountOfIssuers(lngYr)
  Call SQLYr_CountOfSecType(lngYr)
  Call SQLYr_CountOfSecuritiesPerDeal(lngYr)
  Call SQLYr_CountOfSourceType(lngYr)
  '---
  Call SQLYr_AverSizePerDealType(lngYr)
  Call SQLYr_AverSizePerSecType(lngYr)
  '---
  Call CreateUNSourcesSummary(lngYr)
  '---
  Call SQLYr_ShareOfDealType(lngYr)
  Call SQLYr_ShareOfSecType(lngYr)
  Call SQLYr_ShareOfSourceType(lngYr)
  '---
  Call SQLYr_ShareOfLeverageLevel(lngYr)
  Call SQLYr_ShareOfCoverageLevel(lngYr)
  '---
  Call SQLYr_UNClosedVsReviewedDealsPerMonth(lngYr)
  '---
  Call SQLYr_AggregInvAndAverSize(lngYr)
  Call SQLYr_AverSizePerDealType(lngYr)
  Call SQLYr_AverSizePerSecType(lngYr)
  '---
  If lngYr > 1998 Then Call SQLYr_UNPrevYearDealSizeComp(lngYr)
  
  Debug.Print "CreateYearlyQueries: All yearly queries created for " & lngYr
  
  If (MsgBox("Do you want to create quarterly queries for " & lngYr & "?", vbQuestion + vbYesNo, _
             "Create Quarterly Queries?")) = vbYes Then
    For Q = 1 To 4
      Call CreateQuarterlyQueries(lngYr, Q)
    Next Q
    For Q = 1 To 4
      Call CreateLTMQxQueries(lngYr, Q)
    Next Q
  End If

End Function

Function CreateLTMQxQueries(lngYr As Long, iQ As Integer)
  If Not blnOverwrite Then  'ask
    blnOverwrite = (vbYes Mod (MsgBox("Do you want to overwrite existing Q" & iQ & " LTM queries?", _
                                    vbQuestion + vbYesNo, "Overwrite LTM Quarterly Queries?")) = 0)
  End If
  
  Call SQLQr_LTMQEndDealData(lngYr, iQ)
  Call SQLQr_LTMQEndTotReviewedDeals(lngYr, iQ)
  Call SQLQr_LTMQEndReviewedDealsPerMonth(lngYr, iQ)

  Call SQLQr_LTMQEndClosedDealData(lngYr, iQ)
  Call SQLQr_LTMQTotClosedDeals(lngYr, iQ)
  Call SQLQr_LTMQEndClosedDealsPerMonth(lngYr, iQ)

  Call SQLQr_LTMQEndAggregInvAndAverSize(lngYr, iQ)
  Call SQLQr_LTMQEndClosedVsReviewedDeals(lngYr, iQ)
  
  Call SQLQr_LTMQEndCountOfSourceType(lngYr, iQ)
  Call SQLQr_LTMQEndShareOfSourceType(lngYr, iQ)
      
  Debug.Print "CreateLTMQxQueries: All " & lngYr & "-LTMQ" & iQ & " queries created."

End Function

Function SQLYr_AggregInvAndAverSize(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "DealData"
  strSubQry2 = "qry" & lngYr & "TotReviewedDeals"
    
  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry2 & ".ReviewedDeals AS Transactions, Sum([" & _
  cstNoWarrantsQ & "]![sglAmtOffered]) AS AggregateInvPotential, Sum([" & cstNoWarrantsQ & "]" & _
  "![sglAmtOffered])/[" & strSubQry2 & "]![ReviewedDeals] AS AverageSize " & _
  "FROM " & strSubQry2 & ", " & strSubQry1 & " INNER JOIN " & cstNoWarrantsQ & " ON " & _
  strSubQry1 & ".lngDealNum = " & cstNoWarrantsQ & ".lngSecDealNum " & _
  "GROUP BY " & strSubQry1 & ".SelectedYear, " & strSubQry2 & ".ReviewedDeals;"

  strQry = "qry" & lngYr & "AggregInvAndAverSize"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  
End Function

Function SQLYr_ClosedDealData(lngYr As Long)
  Dim strSQL As String, strQry As String

  strSQL = "SELECT tblDeal.*, Year([dteDealDateDisp]) AS YearClosed, DatePart(""q"",[tblDeal]![dteDealDateDisp]) " & _
  "AS QClosed, Year([dteDealDateIn]) AS SelectedYear, tblSource.txtSourceName AS Source, " & _
  "tblIssuer.txtIssuerName AS Issuer, tlkpDealType.txtDealTypeDesc AS DealTypeDesc " & _
  "FROM tlkpDealType RIGHT JOIN (tblIssuer LEFT JOIN (tblSource RIGHT JOIN tblDeal ON " & _
  "tblSource.lngSourceNum=tblDeal.lngDealSourceNum) ON tblIssuer.lngIssuerNum = " & _
  "tblDeal.lngDealIssuerNum) ON tlkpDealType.lngDealTypeNum=tblDeal.lngDealTypeNum " & _
  "WHERE (((Year([dteDealDateDisp]))=" & lngYr & ") AND((tblDeal.lngDealDispNum)=1));"

  strQry = "qry" & lngYr & "ClosedDealData"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  
End Function

Function SQL_ClosedDealData()
  Dim strSQL As String

  strSQL = "SELECT Year([dteDealDateDisp]) AS YearClosed, DatePart(""q"",[tblDeal]![dteDealDateDisp]) " & _
  "AS QClosed, Year([dteDealDateIn]) AS SelectedYear, tblDeal.*, [tblSource].[txtSourceName] AS Source " & _
  "FROM tblSource RIGHT JOIN tblDeal ON [tblSource].[lngSourceNum]=[tblDeal].[lngDealSourceNum] " & _
  "WHERE ((([tblDeal].[lngDealDispNum])=1));"

  Call CreateDBQuery("qryClosedDealData", strSQL, cstDescAll, blnOverwrite)
  
End Function

Function SQL_IssuersFinData()
  Dim strSQL As String
  
  strSQL = "SELECT Year([tblDeal]![dteDealDateIn]) AS SelectedYear, Format([tblDeal]![dteDealDateIn],""q"")"
  strSQL = strSQL & " AS SelectedQ, tblDeal.lngDealNum, tblIssuer.txtIssuerName AS Issuer, "
  strSQL = strSQL & "CLng([tblDeal]![lngDealIssuerNum]) AS IssuerID, IIf([tblFinstat]![lngFinDealNum]<>"
  strSQL = strSQL & "[tblDeal]![lngDealNum],0,[tblFinstat]![sglFinIntExp]) AS FinIntExp, IIf([tblFinstat]!"
  strSQL = strSQL & "[lngFinDealNum]<>[tblDeal]![lngDealNum],0,[tblFinstat]![sglFinTotDebt]) AS FinTotDebt, "
  strSQL = strSQL & "IIf([tblFinstat]![lngFinDealNum]<>[tblDeal]![lngDealNum],0,[tblFinstat]![sglFinEBITDA]) AS "
  strSQL = strSQL & "FinEBITDA, IIf((IsNull([FinIntExp]) Or [FinIntExp]=0),-1,[FinEBITDA]/[FinIntExp]) AS IntCov, "
  strSQL = strSQL & "IIf((IsNull([FinEBITDA]) Or [FinEBITDA]=0),-1,[FinTotDebt]/[FinEBITDA]) AS Lev "
  strSQL = strSQL & "FROM (tblIssuer RIGHT JOIN tblDeal ON tblIssuer.lngIssuerNum = tblDeal.lngDealIssuerNum) "
  strSQL = strSQL & "LEFT JOIN tblFinstat ON tblDeal.lngDealNum = tblFinstat.lngFinDealNum ORDER BY "
  strSQL = strSQL & "Year([tblDeal]![dteDealDateIn]) DESC , Format([tblDeal]![dteDealDateIn],""q"") DESC , " & _
                    "tblDeal.lngDealNum DESC;"

  Call CreateDBQuery("qryIssuersFinData", strSQL, cstDescAll, blnOverwrite)
  
End Function

Function SQL_IssuersWithFinData()
  Dim strSQL As String
  
  strSQL = "SELECT qryIssuersFinData.* FROM qryIssuersFinData WHERE ((([qryIssuersFinData].[IntCov])>-1)) " & _
          "Or ((([qryIssuersFinData].[Lev])>-1));"
  Call CreateDBQuery("qryIssuersWithFinData", strSQL, cstDescAll, blnOverwrite)
  
End Function

Function SQL_InvestmentAmtNonAddOnDeals()
  Dim strSQL As String
  
  strSQL = "SELECT qryClosedDealData.lngDealNum, Sum(qryNonWarrantSecuritiesData.sglSecInvestSize) " & _
    "AS Investment FROM qryClosedDealData LEFT JOIN qryNonWarrantSecuritiesData ON qryClosedDealData." & _
    "lngDealNum = qryNonWarrantSecuritiesData.lngSecDealNum GROUP BY qryClosedDealData.lngDealNum, " & _
    "qryClosedDealData.lngDealTypeNum HAVING (((qryClosedDealData.lngDealTypeNum)<>17));"
 
  Call CreateDBQuery("qryInvestmentAmt-NonAddOnDeals", strSQL, cstDescAll, blnOverwrite)
  
End Function

Function SQL_SourcesWithDealID()
  Dim strSQL As String
 
  strSQL = "SELECT DISTINCT Year([dteDealDateIn]) AS SelectedYear, DatePart(""q"",[dteDealDateIn]) AS " & _
           "SelectedQ, DatePart(""q"",[dteDealDateDisp]) AS QDisposed, tblSource.*, tblDeal.lngDealNum " & _
           "FROM tblSource LEFT JOIN tblDeal ON tblSource.lngSourceNum = tblDeal.lngDealSourceNum " & _
           "WHERE (((tblDeal.lngDealNum) Is Not Null));"
  Call CreateDBQuery(cstSourcesWithDeals, strSQL, cstDescAll, blnOverwrite)
End Function

Function SQL_ClosedDealsSummaryAllMTBL()
  Dim strSQL As String
  Dim strSubQry1 As String, strSubQry2 As String, strINTOQry As String

  strSubQry1 = "qryClosedDealData"
  strSubQry2 = "qryIssuersWithFinData"
  strINTOQry = cstClosedDealsSummAll

  strSQL = "SELECT DISTINCT " & strSubQry1 & ".YearClosed, " & strSubQry1 & ".QClosed, " & strSubQry1 & _
      ".lngDealNum, " & strSubQry1 & ".Issuer, " & strSubQry1 & ".Source," & strSubQry1 & ".DealTypeDesc," & _
      "tlkpRoleType.txtRoleDesc AS [Role Type], IIf([IntCov]<0,""n/a"",Format([IntCov],""#.00"")) AS Coverage," & _
      "IIf([Lev]<0,""n/a"",Format([Lev],""Fixed"")) AS Leverage, tlkpAnalysts.txtAnalLast AS Analyst," & _
      "tlkpAnalysts_1.txtAnalLast AS Analyst2, GetSecTypesStr([" & strSubQry1 & "]![lngDealNum]) " & _
      "AS Securities INTO [" & strINTOQry & "] FROM (((" & strSubQry1 & " LEFT JOIN " & strSubQry2 & _
      " ON " & strSubQry1 & ".lngDealIssuerNum= " & strSubQry2 & ".IssuerID) INNER JOIN tlkpRoleType ON " & _
      strSubQry1 & ".lngRoleType =tlkpRoleType.lngRoleTypeNum) INNER JOIN tlkpAnalysts ON " & strSubQry1 & _
      ".lngDealAnalNum =tlkpAnalysts.lngAnalNum) INNER JOIN tlkpAnalysts AS tlkpAnalysts_1 ON " & strSubQry1 & _
      ".lngDealSecondAnalNum = tlkpAnalysts_1.lngAnalNum ORDER BY " & strSubQry1 & ".YearClosed DESC ," & _
      strSubQry1 & ".QClosed DESC , " & strSubQry1 & ".lngDealNum DESC ," & strSubQry1 & ".Source;"
  
  Call CreateDBQuery(cstClosedDealsSummAllMTBL, strSQL, cstDescAll, blnOverwrite)
  Call RunDBQuery(cstClosedDealsSummMTBL)
End Function

Function SQL_ClosedDealsSummary()
  Dim strSQL As String, strINTOQry As String
  
  strINTOQry = cstClosedDealsSummAll
  On Error Resume Next
start:
  CurrentDb.Execute cstClosedDealsSummAllMTBL, dbSeeChanges 'dbFailOnError
  If Err <> 0 Then
    If Err = 3010 Then
      Err.Clear
      With CurrentDb.TableDefs
        .Delete cstClosedDealsSummAll
        .Refresh
      End With
      Resume start:
    Else
      GoTo SQL_ClosedDealsSummaryErr
    End If
  End If
  
  strSQL = "SELECT tblClosedDealsLegacySummary.[Yr Closed],tblClosedDealsLegacySummary.[Q Closed]," & _
    "tblClosedDealsLegacySummary.Issuer, tblClosedDealsLegacySummary.Industry," & _
    "tblClosedDealsLegacySummary.[Deal Type],tblClosedDealsLegacySummary.Source," & _
    "tblClosedDealsLegacySummary.[Source Type],tblClosedDealsLegacySummary.[Role Type]," & _
    "tblClosedDealsLegacySummary.Coverage, tblClosedDealsLegacySummary.Leverage, " & _
    "tblClosedDealsLegacySummary.Investment,tblClosedDealsLegacySummary.Securities, " & _
    "tblClosedDealsLegacySummary.Analyst FROM tblClosedDealsLegacySummary " & _
    "UNION SELECT [" & strINTOQry & "].[Yr Closed], [" & strINTOQry & "].[Q Closed], [" & _
    strINTOQry & "].Issuer, [" & strINTOQry & "].Industry, [" & strINTOQry & "].[Deal Type], [" & _
    strINTOQry & "].Source, [" & strINTOQry & "].[Source Type],[" & strINTOQry & "].[Role Type], [" & _
    strINTOQry & "].Coverage,[" & strINTOQry & "].Leverage, [" & strINTOQry & "].Investment, [" & _
    strINTOQry & "].Securities,[" & strINTOQry & "].Analyst FROM [" & strINTOQry & "] " & _
    "ORDER BY tblClosedDealsLegacySummary.[Yr Closed] DESC , tblClosedDealsLegacySummary.[Q Closed] DESC;"

  Call CreateDBQuery("qryClosedDealsSummary", strSQL, cstDescAll, blnOverwrite)
  
SQL_ClosedDealsSummaryExit:
  Exit Function
   
SQL_ClosedDealsSummaryErr:
  MsgBox "Error: " & Err & vbCrLf & "Desc:  " & Err.Description, , "SQL_ClosedDealsSummary"
  Resume SQL_ClosedDealsSummaryExit
End Function

Function SQLYr_ClosedDealsPerMonth(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
   
  strSubQry1 = "qry" & lngYr & "ClosedDealData"

  strSQL = "SELECT " & strSubQry1 & ".YearClosed, Format([" & strSubQry1 & _
           "]![dteDealDateDisp],""mmm"") AS SelectedMonth, Count(" & strSubQry1 & _
           ".lngDealNum) AS ClosedDeals, True AS Closed, Month([" & strSubQry1 & _
           "]![dteDealDateDisp]) AS MonthNum FROM " & strSubQry1 & " GROUP BY " & strSubQry1 & _
           ".YearClosed, Format([" & strSubQry1 & "]![dteDealDateDisp],""mmm""),True," & "Month([" & _
           strSubQry1 & "]![dteDealDateDisp]) ORDER BY Month([" & strSubQry1 & "]![dteDealDateDisp]);"

  strQry = "qry" & lngYr & "ClosedDealsPerMonth"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  
End Function

Function SQLYr_ClosedDealsSummary(lngYr As Long)
  Dim strSQL As String, strQry As String, strSourceTbl As String
  
  strSourceTbl = "tblClosedDealsSummary-Yr" & lngYr
  
  strSQL = "SELECT DISTINCTROW [" & strSourceTbl & "].YearClosed, [" & strSourceTbl & "].QClosed, " & _
  "[" & strSourceTbl & "].Analyst, [" & strSourceTbl & "].Analyst2, [" & strSourceTbl & "].Issuer, " & _
  "[" & strSourceTbl & "].Source, [" & strSourceTbl & "].DealTypeDesc, " & _
  "[" & strSourceTbl & "].[Role Type], [" & strSourceTbl & "].Coverage, " & _
  "[" & strSourceTbl & "].Leverage, GetSecTypesStr([" & strSourceTbl & "]![lngDealNum]) AS " & _
  "Securities FROM [" & strSourceTbl & "] ORDER BY [" & strSourceTbl & "].YearClosed DESC;"

  strQry = "qry" & lngYr & "ClosedDealSummary"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  
End Function

Function SQLYr_ClosedDealsSummaryMTBL(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String, strSourceTbl As String
  
  On Error GoTo ProcErr
  
  strSourceTbl = "tblClosedDealsSummary-Yr" & lngYr

  strSubQry1 = "qry" & lngYr & "ClosedDealData"
  strSubQry2 = "qry" & lngYr & "IssuersWithFinData"

  strSQL = "SELECT DISTINCT " & strSubQry1 & ".YearClosed, " & strSubQry1 & ".QClosed, " & _
    strSubQry1 & ".lngDealNum, " & strSubQry1 & ".Issuer, " & strSubQry1 & ".Source, " & _
    strSubQry1 & ".DealTypeDesc, tlkpRoleType.txtRoleDesc AS [Role Type], " & _
    "IIf([IntCov]<0,""n/a"",Format([IntCov],""#.00"")) AS Coverage, IIf([Lev]<0,""n/a""," & _
    "Format([Lev],""Fixed"")) AS Leverage, tlkpAnalysts.txtAnalLast AS Analyst, " & _
    "tlkpAnalysts_1.txtAnalLast AS Analyst2,GetSecTypesStr(" & strSubQry1 & _
    ".lngDealNum) AS Securities INTO [" & strSourceTbl & "] " & _
    "FROM (((" & strSubQry1 & " LEFT JOIN " & strSubQry2 & " ON " & strSubQry1 & ".lngDealIssuerNum = " & _
    strSubQry2 & ".IssuerID) INNER JOIN tlkpRoleType ON " & strSubQry1 & ".lngRoleType = " & _
    "tlkpRoleType.lngRoleTypeNum) INNER JOIN tlkpAnalysts ON " & strSubQry1 & ".lngDealAnalNum = " & _
    "tlkpAnalysts.lngAnalNum) INNER JOIN tlkpAnalysts AS tlkpAnalysts_1 ON " & strSubQry1 & _
    ".lngDealSecondAnalNum = tlkpAnalysts_1.lngAnalNum ORDER BY  " & strSubQry1 & _
    ".YearClosed DESC , " & strSubQry1 & ".QClosed DESC , " & strSubQry1 & ".lngDealNum DESC , " & _
    strSubQry1 & ".Source;"
  
  strQry = "qry" & lngYr & "ClosedDealSummaryMTBL"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  Call RunDBQuery(strQry)

ProcExit:
  Exit Function
ProcErr:
  If Err = 3010 Then 'INTO tbl [tblClosedDealsSummary-Yr & lngYr] already exists
    Err.Clear
    Call DeleteDBQry(strSourceTbl)
    Call RunDBQuery(strQry)
  Else
    MsgBox "Error: " & Err.Number & " : " & Err.Description, vbExclamation, "RunDBQuery"
  End If
  Resume ProcExit
End Function

Function SQLYr_ClosedDealsSummaryWithSecFinStat(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "ClosedDealData"
  strSubQry2 = "qry" & lngYr & "IssuersFinData"

  strSQL = "SELECT DISTINCT " & strSubQry1 & ".YearClosed, " & strSubQry1 & ".QClosed,  " & _
  strSubQry2 & ".Issuer, tblSource.txtSourceName AS Source, tlkpDealType.txtDealTypeDesc AS [Deal Type], " & _
  strSubQry2 & ".IntCov AS Coverage, " & strSubQry2 & ".Lev AS Leverage, " & cstNoWarrantsQ & "." & _
  "lngSecTypeNum, " & strSubQry1 & ".lngDealNum " & _
  "FROM (((" & strSubQry1 & " INNER JOIN " & strSubQry2 & " ON " & strSubQry1 & ".lngDealNum = " & _
  strSubQry2 & ".lngDealNum) LEFT JOIN tblSource ON " & strSubQry1 & ".lngDealSourceNum = " & _
  "tblSource.lngSourceNum) INNER JOIN tlkpDealType ON " & strSubQry1 & ".lngDealTypeNum = " & _
  "tlkpDealType.lngDealTypeNum) LEFT JOIN " & cstNoWarrantsQ & " ON " & strSubQry1 & ".lngDealNum = " & _
  cstNoWarrantsQ & ".lngSecDealNum ORDER BY " & strSubQry1 & ".QClosed DESC, " & _
  cstNoWarrantsQ & ".lngSecTypeNum DESC , " & strSubQry1 & ".lngDealNum DESC;"

  strQry = "qry" & lngYr & "ClosedDealsSummaryWithSecFinStat"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  
End Function

Function SQLYr_CountOfDealType(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "DealData"    'Xqry2001DealData

  strSQL = "SELECT DISTINCT Count(" & strSubQry1 & ".lngDealNum) AS [Deal Type Count], " & _
  "tlkpDealType.txtDealTypeDesc AS [Deal Type] FROM " & strSubQry1 & " INNER JOIN tlkpDealType ON " & _
  strSubQry1 & ".lngDealTypeNum = tlkpDealType.lngDealTypeNum GROUP BY tlkpDealType.txtDealTypeDesc " & _
  "HAVING (((Count( " & strSubQry1 & ".lngDealNum)) > 0)) ORDER BY Count(" & strSubQry1 & ".lngDealNum) DESC;"
  
  strQry = "qry" & lngYr & "CountOfDealType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)

End Function

Function SQLYr_CountOfIssuers(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "IssuersFinData"    'Xqry2001IssuersFinData

  strSQL = "SELECT [ " & strSubQry1 & "].[SelectedYear], Count([ " & strSubQry1 & "].[lngDealIssuerNum]) " & _
  "AS [Issuers Count] FROM  " & strSubQry1 & " GROUP BY [ " & strSubQry1 & "].[SelectedYear];"

  strQry = "qry" & lngYr & "CountOfIssuers"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  
End Function

Function SQLYr_CountOfIssuersWithFinData(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "IssuersWithFinData"
  
  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, Count(" & strSubQry1 & ".lngDealIssuerNum) " & _
          "AS [Issuers Count] FROM  " & strSubQry1 & " GROUP BY " & strSubQry1 & ".SelectedYear;"

  strQry = "qry" & lngYr & "CountOfIssuersWithFinData"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  
End Function

Function SQLYr_CountOfSecType(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "DealData"

  strSQL = "SELECT DISTINCT Count(" & strSubQry1 & ".lngDealNum) AS [Sec Type Count], " & _
          "tlkpSecType.txtSecType AS [Sec Type] FROM " & strSubQry1 & " INNER JOIN (tlkpSecType " & _
          "INNER JOIN tblSecurity ON tlkpSecType.lngSecTypeIdx=tblSecurity.lngSecTypeNum) ON " & _
          strSubQry1 & ".lngDealNum=tblSecurity.lngSecDealNum GROUP BY tlkpSecType.txtSecType " & _
          "HAVING ((Count( " & strSubQry1 & ".lngDealNum)) > 0) ORDER BY Count(" & strSubQry1 & _
          ".lngDealNum) DESC;"
  
  strQry = "qry" & lngYr & "CountOfSecType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  
End Function

Function SQLYr_CountOfSecuritiesPerDeal(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "DealData"    'Xqry2001DealData
  'cstNoWarrantsQ  = " & cstNoWarrantsQ & "

  strSQL = "SELECT  " & strSubQry1 & ".SelectedYear,  " & strSubQry1 & ".lngDealIssuerNum,  " & strSubQry1 & ".lngDealNum, " & _
  "Count( " & cstNoWarrantsQ & ".lngSecNum) AS [Securities Count] FROM  " & strSubQry1 & " LEFT JOIN " & _
  cstNoWarrantsQ & " ON  " & strSubQry1 & ".lngDealNum =  " & cstNoWarrantsQ & ".lngSecDealNum " & _
  "GROUP BY  " & strSubQry1 & ".SelectedYear,  " & strSubQry1 & ".lngDealIssuerNum,  " & strSubQry1 & ".lngDealNum;"
  
  strQry = "qry" & lngYr & "CountOfSecuritiesPerDeal"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)

End Function

Function SQLYr_CountOfSourceType(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "DealData"
  strSQL = "SELECT DISTINCT tlkpSourceType.txtSourceTypeDesc, Count(" & strSubQry1 & _
    ".lngDealNum) AS [Deal Count] FROM (tblSource INNER JOIN tlkpSourceType ON " & _
    "tblSource.lngSourceTypeNum = tlkpSourceType.lngSourceTypeNum) INNER JOIN " & _
    strSubQry1 & " ON tblSource.lngSourceNum = " & strSubQry1 & ".lngDealSourceNum " & _
    "GROUP BY tlkpSourceType.txtSourceTypeDesc HAVING ((Count(" & strSubQry1 & _
    ".lngDealNum)) > 0) ORDER BY Count(" & strSubQry1 & ".lngDealNum) DESC;"

  strQry = "qry" & lngYr & "CountOfSourceType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLYr_DealData(lngYr As Long)
  Dim strSQL As String, strQry As String
  
  strSQL = "SELECT Year([dteDealDateIn]) AS SelectedYear, DatePart(""q"",[dteDealDateIn]) AS SelectedQ, " & _
  "DatePart(""q"",[dteDealDateDisp]) AS QDisposed, tblDeal.* FROM tblDeal WHERE (((Year([dteDealDateIn])) = " & _
  lngYr & ")) ORDER BY [tblDeal].[dteDealDateIn] DESC;"
  
  strQry = "qry" & lngYr & "DealData"
   Call CreateDBQuery(strQry, strSQL, cstDescAll, blnOverwrite)
End Function

Function SQLYr_IssuerCoverage(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "IssuersWithFinData"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, Count(" & strSubQry1 & ".lngDealIssuerNum) " & _
    "AS Issuers, IIf([IntCov]<=0,""No data"",IIf([IntCov]<=1,""]0, 1]"", IIf([IntCov]<=2,""]1, 2]""," & _
    "IIf([IntCov]<=3,""]2, 3]"",IIf([IntCov]<=4,""]3, 4]"",IIf([IntCov]<=5,""]4, 5]""," & _
    "IIf([IntCov]<=6,""]5, 6]"",IIf([IntCov]<=7,""]6, 7]"","">7"")))))))) AS [Group] FROM " & strSubQry1 & _
    " GROUP BY " & strSubQry1 & ".SelectedYear, IIf([IntCov]<=0,""No data"",IIf([IntCov]<=1,""]0, 1]""," & _
    "IIf([IntCov]<=2,""]1, 2]"",IIf([IntCov]<=3,""]2, 3]"",IIf([IntCov]<=4,""]3, 4]""," & _
    "IIf([IntCov]<=5,""]4, 5]"",IIf([IntCov]<=6,""]5, 6]"",IIf([IntCov]<=7,""]6, 7]"","">7"")))))))) " & _
    "ORDER BY IIf([IntCov]<=0,""No data"",IIf([IntCov]<=1,""]0, 1]"",IIf([IntCov]<=2,""]1, 2]""," & _
    "IIf([IntCov]<=3,""]2, 3]"",IIf([IntCov]<=4,""]3, 4]"",IIf([IntCov]<=5,""]4, 5]""," & _
    "IIf([IntCov]<=6,""]5, 6]"",IIf([IntCov]<=7,""]6, 7]"","">7""))))))));"
  
  strQry = "qry" & lngYr & "IssuerCoverage"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLQr_IssuerCoverage(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String, strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "IssuersWithFinData"
  
  strSQL = "SELECT " & strSubQry1 & ".SelectedYear," & strSubQry1 & ".SelectedQ, Count(" & strSubQry1 & _
      ".lngDealIssuerNum) AS Issuers, IIf([IntCov]<=0,""No data"",IIf([IntCov]<=1,""]0, 1]""," & _
      "IIf([IntCov]<=2,""]1, 2]"", IIf([IntCov]<=3,""]2, 3]"",IIf([IntCov]<=4,""]3, 4]""," & _
      "IIf([IntCov]<=5,""]4, 5]"",IIf([IntCov]<=6,""]5, 6]"",IIf([IntCov]<=7,""]6, 7]"","">7"")))))))) " & _
      "AS [Group] FROM " & strSubQry1 & " GROUP BY " & strSubQry1 & ".SelectedYear," & strSubQry1 & _
      ".SelectedQ, IIf([IntCov]<=0,""No data"",IIf([IntCov]<=1,""]0, 1]"",IIf([IntCov]<=2,""]1, 2]""," & _
      "IIf([IntCov]<=3,""]2, 3]"",IIf([IntCov]<=4,""]3, 4]"",IIf([IntCov]<=5,""]4, 5]""," & _
      "IIf([IntCov]<=6,""]5, 6]"",IIf([IntCov]<=7,""]6, 7]"","">7"")))))))) ORDER BY " & _
      "IIf([IntCov]<=0,""No data"",IIf([IntCov]<=1,""]0, 1]"",IIf([IntCov]<=2,""]1, 2]""," & _
      "IIf([IntCov]<=3,""]2, 3]"",IIf([IntCov]<=4,""]3, 4]"",IIf([IntCov]<=5,""]4, 5]""," & _
      "IIf([IntCov]<=6,""]5, 6]"",IIf([IntCov]<=7,""]6, 7]"","">7""))))))));"
  
  strQry = "qry" & lngYr & "Q" & iQtr & "IssuerCoverage"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_IssuerLeverage(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "IssuersWithFinData"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear," & strSubQry1 & ".SelectedQ, Count(" & strSubQry1 & _
      ".lngDealIssuerNum) AS Issuers, IIf([Lev]<=0,""No data"",IIf([Lev]<=1,""]0, 1]""," & _
      "IIf([Lev]<=2,""]1, 2]"", IIf([Lev]<=3,""]2, 3]"",IIf([Lev]<=4,""]3, 4]"",IIf([Lev]<=5,""]4, 5]""," & _
      "IIf([Lev]<=6,""]5, 6]"",IIf([Lev]<=7,""]6, 7]"","">7"")))))))) AS [Group] FROM " & strSubQry1 & _
      " GROUP BY " & strSubQry1 & ".SelectedYear," & strSubQry1 & ".SelectedQ, IIf([Lev]<=0,""No data""," & _
      "IIf([Lev]<=1,""]0, 1]"",IIf([Lev]<=2,""]1, 2]"",IIf([Lev]<=3,""]2, 3]"",IIf([Lev]<=4,""]3, 4]""," & _
      "IIf([Lev]<=5,""]4, 5]"",IIf([Lev]<=6,""]5, 6]"",IIf([Lev]<=7,""]6, 7]"","">7"")))))))) " & _
      "ORDER BY IIf([Lev]<=0,""No data"",IIf([Lev]<=1,""]0, 1]"",IIf([Lev]<=2,""]1, 2]""," & _
      "IIf([Lev]<=3,""]2, 3]"",IIf([Lev]<=4,""]3, 4]"",IIf([Lev]<=5,""]4, 5]""," & _
      "IIf([Lev]<=6,""]5, 6]"",IIf([Lev]<=7,""]6, 7]"","">7""))))))));"
  
  strQry = "qry" & lngYr & "Q" & iQtr & "IssuerLeverage"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLYr_IssuerLeverage(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "IssuersWithFinData"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, Count(" & strSubQry1 & ".lngDealIssuerNum) " & _
    "AS Issuers, IIf([Lev]<=0,""No data"",IIf([Lev]<=1,""]0, 1]"", IIf([Lev]<=2,""]1, 2]""," & _
    "IIf([Lev]<=3,""]2, 3]"",IIf([Lev]<=4,""]3, 4]"",IIf([Lev]<=5,""]4, 5]""," & _
    "IIf([Lev]<=6,""]5, 6]"",IIf([Lev]<=7,""]6, 7]"","">7"")))))))) AS [Group] FROM " & strSubQry1 & _
    " GROUP BY " & strSubQry1 & ".SelectedYear, IIf([Lev]<=0,""No data"",IIf([Lev]<=1,""]0, 1]""," & _
    "IIf([Lev]<=2,""]1, 2]"",IIf([Lev]<=3,""]2, 3]"",IIf([Lev]<=4,""]3, 4]""," & _
    "IIf([Lev]<=5,""]4, 5]"",IIf([Lev]<=6,""]5, 6]"",IIf([Lev]<=7,""]6, 7]"","">7"")))))))) " & _
    "ORDER BY IIf([Lev]<=0,""No data"",IIf([Lev]<=1,""]0, 1]"",IIf([Lev]<=2,""]1, 2]""," & _
    "IIf([Lev]<=3,""]2, 3]"",IIf([Lev]<=4,""]3, 4]"",IIf([Lev]<=5,""]4, 5]""," & _
    "IIf([Lev]<=6,""]5, 6]"",IIf([Lev]<=7,""]6, 7]"","">7""))))))));"
  
  strQry = "qry" & lngYr & "IssuerLeverage"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLYr_IssuersFinData(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "DealData"
  
  strSQL = "SELECT " & strSubQry1 & ".lngDealNum, " & strSubQry1 & ".SelectedYear, " & strSubQry1 & _
    ".SelectedQ, tblIssuer.txtIssuerName AS Issuer, " & strSubQry1 & ".lngDealIssuerNum, " & _
    "IIf([tblFinstat]![lngFinDealNum]<>[tblDeal]![lngDealNum],0, " & _
    "[tblFinstat]![sglFinIntExp]) AS FinIntExp, IIf([tblFinstat]![lngFinDealNum]<>[tblDeal]![lngDealNum],0, " & _
    "[tblFinstat]![sglFinTotDebt]) AS FinTotDebt, IIf([tblFinstat]![lngFinDealNum]<>[tblDeal]![lngDealNum],0, " & _
    "[tblFinstat]![sglFinEBITDA]) AS FinEBITDA, IIf((IsNull([FinIntExp]) Or [FinIntExp]=0),-1,[FinEBITDA]/[FinIntExp]) " & _
    "AS IntCov, IIf((IsNull([FinEBITDA]) Or [FinEBITDA]=0),-1,[FinTotDebt]/[FinEBITDA]) AS Lev " & _
    "FROM " & strSubQry1 & " INNER JOIN (tblIssuer LEFT JOIN tblFinstat ON tblIssuer.lngIssuerNum= " & _
    "tblFinstat.lngFinIssuerNum) ON " & strSubQry1 & ".lngDealIssuerNum=tblIssuer.lngIssuerNum " & _
    "WHERE (((IIf([tblFinstat]![lngFinDealNum]<>[tblDeal]![lngDealNum],0,[tblFinstat]![sglFinIntExp])) Is Not Null) " & _
    "AND ((IIf([tblFinstat]![lngFinDealNum]<>[tblDeal]![lngDealNum],0,[tblFinstat]![sglFinTotDebt])) Is Not Null) " & _
    "AND ((IIf([tblFinstat]![lngFinDealNum]<>[tblDeal]![lngDealNum],0,[tblFinstat]![sglFinEBITDA])) Is Not Null));"

  strQry = "qry" & lngYr & "IssuersFinData"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  
End Function

Function SQLQr_IssuersFinData(lngYr As Long, iQtr As Integer)  'subqry to QxIssuersWithFinData
  Dim strSQL As String, strQry As String, strSubQry As String
  
  strSubQry = "qry" & lngYr & "Q" & iQtr & "DealData"    'Xqry2001Q1DealData
 
  strSQL = "SELECT " & strSubQry & ".lngDealNum, " & strSubQry & ".SelectedYear, " & strSubQry
  strSQL = strSQL & ".SelectedQ, tblIssuer.txtIssuerName AS Issuer, " & strSubQry
  strSQL = strSQL & ".lngDealIssuerNum, IIf([tblFinstat]![lngFinDealNum]"
  strSQL = strSQL & "<>[tblDeal]![lngDealNum],0,[tblFinstat]![sglFinIntExp]) AS FinIntExp, "
  strSQL = strSQL & "IIf([tblFinstat]![lngFinDealNum]<>[tblDeal]![lngDealNum],0,[tblFinstat]!"
  strSQL = strSQL & "[sglFinTotDebt]) AS FinTotDebt, IIf([tblFinstat]![lngFinDealNum]<>[tblDeal]!"
  strSQL = strSQL & "[lngDealNum],0,[tblFinstat]![sglFinEBITDA]) AS FinEBITDA, IIf((IsNull([FinIntExp])"
  strSQL = strSQL & " Or [FinIntExp]=0),-1,[FinEBITDA]/[FinIntExp]) AS IntCov, IIf((IsNull([FinEBITDA])"
  strSQL = strSQL & " Or [FinEBITDA]=0),-1,[FinTotDebt]/[FinEBITDA]) AS Lev "
  strSQL = strSQL & "FROM " & strSubQry & " INNER JOIN (tblIssuer LEFT JOIN "
  strSQL = strSQL & "tblFinstat ON tblIssuer.lngIssuerNum = tblFinstat.lngFinIssuerNum) ON "
  strSQL = strSQL & strSubQry & ".lngDealIssuerNum = tblIssuer.lngIssuerNum "
  strSQL = strSQL & "WHERE (((IIf([tblFinstat]![lngFinDealNum]<>[tblDeal]![lngDealNum],0,"
  strSQL = strSQL & "[tblFinstat]![sglFinIntExp])) Is Not Null) AND ((IIf([tblFinstat]!"
  strSQL = strSQL & "[lngFinDealNum]<>[tblDeal]![lngDealNum],0,[tblFinstat]![sglFinTotDebt])) "
  strSQL = strSQL & "Is Not Null) AND ((IIf([tblFinstat]![lngFinDealNum]<>[tblDeal]![lngDealNum],0,"
  strSQL = strSQL & "[tblFinstat]![sglFinEBITDA])) Is Not Null));"

  strQry = "qry" & lngYr & "Q" & iQtr & "IssuersFinData"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
  
End Function

Function SQLQr_CountOfIssuers(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String

  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "DealData"     'Xqry2001Q1DealData

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, Count(tblIssuer.txtIssuerName) " & _
    "AS Issuers FROM " & strSubQry1 & " INNER JOIN tblIssuer ON " & strSubQry1 & ".lngDealIssuerNum = " & _
    "tblIssuer.lngIssuerNum GROUP BY " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ;"

  strQry = "qry" & lngYr & "Q" & iQtr & "CountOfIssuers"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_CountOfSourceType(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String

  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "DealData"     'Xqry2001Q1DealData

  strSQL = "SELECT DISTINCT Count(" & strSubQry1 & ".lngDealNum) AS [Deal Count], " & _
    "tlkpSourceType.txtSourceTypeDesc FROM " & strSubQry1 & " INNER JOIN (tblSource INNER JOIN tlkpSourceType " & _
    "ON tblSource.lngSourceTypeNum = tlkpSourceType.lngSourceTypeNum) ON " & strSubQry1 & _
    ".lngDealSourceNum = tblSource.lngSourceNum GROUP BY tlkpSourceType.txtSourceTypeDesc " & _
    "HAVING (((Count(" & strSubQry1 & ".lngDealNum)) > 0)) ORDER BY Count(" & strSubQry1 & ".lngDealNum) DESC;"
  'Debug.Print "QxCountOfSourceType sql: " & strSQL

  strQry = "qry" & lngYr & "Q" & iQtr & "CountOfSourceType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_CountOfSecType(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String, strSubQry1 As String
  strSQL = ""
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "DealData"     'qryYrxQxDealData
  
  strSQL = "SELECT DISTINCT Count(" & strSubQry1 & ".lngDealNum) AS [Sec Type Count], " & _
          "tlkpSecType.txtSecType AS [Sec Type] FROM " & strSubQry1 & " INNER JOIN " & _
          "(tlkpSecType INNER JOIN tblSecurity ON tlkpSecType.lngSecTypeIdx = " & _
          "tblSecurity.lngSecTypeNum) ON " & strSubQry1 & ".lngDealNum = tblSecurity.lngSecDealNum " & _
          "GROUP BY tlkpSecType.txtSecType HAVING (((Count( " & strSubQry1 & ".lngDealNum)) > 0)) " & _
          "ORDER BY Count(" & strSubQry1 & ".lngDealNum) DESC;"
  'Debug.Print strSQL
  strQry = "qry" & lngYr & "Q" & iQtr & "CountOfSecType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
  
End Function

Function SQLQr_CountOfDealType(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String, strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "DealData"     'qry200xQxDealData

  strSQL = "SELECT DISTINCT Count(" & strSubQry1 & ".lngDealNum) AS [Deal Type Count], " & _
  "tlkpDealType.txtDealTypeDesc AS [Deal Type] FROM " & strSubQry1 & " INNER JOIN tlkpDealType ON " & _
  strSubQry1 & ".lngDealTypeNum = tlkpDealType.lngDealTypeNum GROUP BY tlkpDealType.txtDealTypeDesc " & _
  "HAVING (((Count( " & strSubQry1 & ".lngDealNum)) > 0)) ORDER BY Count(" & strSubQry1 & ".lngDealNum) DESC;"

  strQry = "qry" & lngYr & "Q" & iQtr & "CountOfDealType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)

End Function

Function SQLQr_CountOfIssuersWithFinData(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String, strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "IssuersWithFinData"    'qryYrxQxIssuersWithFinData

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, Count(" & _
           strSubQry1 & ".Issuer) AS Issuers FROM " & strSubQry1 & _
           " GROUP BY " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ;"
  strQry = "qry" & lngYr & "Q" & iQtr & "CountOfIssuersWithFinData"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
  
End Function

Function SQLQr_IssuersWithFinData(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String, strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "IssuersFinData"     'XqryQ1IssuersFinData
  strSQL = "SELECT " & strSubQry1 & ".* FROM " & strSubQry1 & _
          " WHERE ((([" & strSubQry1 & "].[IntCov])>-1)) Or ((([" & strSubQry1 & "].[Lev])>-1));"

  strQry = "qry" & lngYr & "Q" & iQtr & "IssuersWithFinData"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)

End Function

Function SQLYr_IssuersWithFinData(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "IssuersFinData"
 
  strSQL = "SELECT " & strSubQry1 & ".* , " & strSubQry1 & ".lngDealIssuerNum AS IssuerID FROM " & _
    strSubQry1 & " WHERE ((([" & strSubQry1 & "].[IntCov])>-1)) Or ((([" & strSubQry1 & "].[Lev])>-1));"

  strQry = "qry" & lngYr & "IssuersWithFinData"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)

End Function

Function SQLYr_AverSizePerDealType(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "CountOfDealType"
  strSubQry2 = "qry" & lngYr & "DealData"
  
  strSQL = "SELECT " & strSubQry1 & ".[Deal Type], " & strSubQry1 & ".[Deal Type Count] AS [Deal Count], " & _
    "Sum([" & cstNoWarrantsQ & "]![sglAmtOffered])/[" & strSubQry1 & "]![Deal Type Count] AS [Average Size] " & _
    "FROM (" & strSubQry1 & " INNER JOIN (tlkpDealType INNER JOIN " & strSubQry2 & " ON " & _
    "tlkpDealType.lngDealTypeNum = " & strSubQry2 & ".lngDealTypeNum) ON " & strSubQry1 & ".[Deal Type] = " & _
    "tlkpDealType.txtDealTypeDesc) LEFT JOIN " & cstNoWarrantsQ & " ON " & strSubQry2 & ".lngDealNum = " & _
    cstNoWarrantsQ & ".lngSecDealNum GROUP BY " & strSubQry1 & ".[Deal Type], " & strSubQry1 & _
    ".[Deal Type Count] ORDER BY " & strSubQry1 & ".[Deal Type Count] DESC;"

  strQry = "qry" & lngYr & "AverSizePerDealType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLQr_AverSizePerDealType(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "CountOfDealType"
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "DealData"

  strSQL = "SELECT " & strSubQry1 & ".[Deal Type], " & strSubQry1 & ".[Deal Type Count] AS [Deal Count], " & _
    "Sum([" & cstNoWarrantsQ & "]![sglAmtOffered])/[" & strSubQry1 & "]![Deal Type Count] AS [Average Size] " & _
    "FROM (" & strSubQry1 & " INNER JOIN (tlkpDealType INNER JOIN " & strSubQry2 & " ON " & _
    "tlkpDealType.lngDealTypeNum = " & strSubQry2 & ".lngDealTypeNum) ON " & strSubQry1 & ".[Deal Type] = " & _
    "tlkpDealType.txtDealTypeDesc) LEFT JOIN " & cstNoWarrantsQ & " ON " & strSubQry2 & ".lngDealNum = " & _
    cstNoWarrantsQ & ".lngSecDealNum GROUP BY " & strSubQry1 & ".[Deal Type], " & strSubQry1 & _
    ".[Deal Type Count] ORDER BY " & strSubQry1 & ".[Deal Type Count] DESC;"

  strQry = "qry" & lngYr & "Q" & iQtr & "AverSizePerDealType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLYr_AverSizePerSecType(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "CountOfSecType"
  strSubQry2 = "qry" & lngYr & "DealData"

  strSQL = "SELECT " & strSubQry1 & ".[Sec Type], " & strSubQry1 & ".[Sec Type Count] AS [Security Count], " & _
    "Sum([" & cstNoWarrantsQ & "]![sglAmtOffered])/[" & strSubQry1 & "]![Sec Type Count] AS [Average Size] " & _
    "FROM (" & strSubQry1 & " INNER JOIN (tlkpSecType INNER JOIN (" & strSubQry2 & " INNER JOIN " & _
    cstNoWarrantsQ & " ON " & strSubQry2 & ".lngDealNum = " & cstNoWarrantsQ & ".lngSecDealNum) ON " & _
    "tlkpSecType.lngSecTypeIdx = " & cstNoWarrantsQ & ".lngSecTypeNum) ON " & strSubQry1 & ".[Sec Type] = " & _
    "tlkpSecType.txtSecType) GROUP BY " & strSubQry1 & ".[Sec Type], " & strSubQry1 & _
    ".[Sec Type Count] ORDER BY " & strSubQry1 & ".[Sec Type Count] DESC;"
        
  strQry = "qry" & lngYr & "AverSizePerSecType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLQr_AverSizePerSecType(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "CountOfSecType"
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "DealData"
  
  strSQL = "SELECT " & strSubQry1 & ".[Sec Type], " & strSubQry1 & ".[Sec Type Count] AS [Security Count], " & _
    "Sum([" & cstNoWarrantsQ & "]![sglAmtOffered])/[" & strSubQry1 & "]![Sec Type Count] AS [Average Size] " & _
    "FROM (" & strSubQry1 & " INNER JOIN (tlkpSecType INNER JOIN (" & strSubQry2 & " INNER JOIN " & _
    cstNoWarrantsQ & " ON " & strSubQry2 & ".lngDealNum = " & cstNoWarrantsQ & ".lngSecDealNum) ON " & _
    "tlkpSecType.lngSecTypeIdx = " & cstNoWarrantsQ & ".lngSecTypeNum) ON " & strSubQry1 & ".[Sec Type] = " & _
    "tlkpSecType.txtSecType) GROUP BY " & strSubQry1 & ".[Sec Type], " & strSubQry1 & _
    ".[Sec Type Count] ORDER BY " & strSubQry1 & ".[Sec Type Count] DESC;"

    
  strQry = "qry" & lngYr & "Q" & iQtr & "AverSizePerSecType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_LTMQEndAggregInvAndAverSize(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "LTM-Q" & iQtr & "TotReviewedDeals"
  strSubQry2 = "qry" & lngYr & "LTM-Q" & iQtr & "EndDealData"

  strSQL = "SELECT [" & strSubQry1 & "].[ReviewedDeals] AS Transactions, " & _
    "Sum(IIf([" & cstNoWarrantsQ & "]![sglAmtOffered]=0,[" & cstNoWarrantsQ & "]![sglTrancheSize], [" & _
    cstNoWarrantsQ & "]![sglAmtOffered])) AS AggregateInvPotential, " & _
    "Sum(IIf([" & cstNoWarrantsQ & "]![sglAmtOffered]=0,[" & cstNoWarrantsQ & "]![sglTrancheSize], [" & _
    cstNoWarrantsQ & "]![sglAmtOffered]))/[" & strSubQry1 & "]![ReviewedDeals] AS AverageSize " & _
    "FROM [" & strSubQry1 & "], " & cstNoWarrantsQ & " INNER JOIN [" & strSubQry2 & "] ON [" & _
    cstNoWarrantsQ & "].[lngSecDealNum]=[" & strSubQry2 & "].[lngDealNum] " & _
    "GROUP BY [" & strSubQry1 & "].[ReviewedDeals];"
  
  strQry = "qry" & lngYr & "LTM-Q" & iQtr & "AggregInvAndAverSize"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_LTMQEndClosedDealData(lngYr As Long, iQtr As Integer) As String ' 42
  Dim strSQL As String, strQry As String
  Dim dte1 As Date, dte2 As Date
  
  Select Case iQtr
    Case 1
      'Between #4/1/00# And #3/31/01#
      dte1 = DateSerial(lngYr - 1, 4, 1)
      dte2 = DateSerial(lngYr, 3, 31)
    Case 2
      'Between #7/1/00# And #6/30/01#
      dte1 = DateSerial(lngYr - 1, 7, 1)
      dte2 = DateSerial(lngYr, 6, 30)
    Case 3
      'Between #11/1/00# And #10/31/01#
      dte1 = DateSerial(lngYr - 1, 11, 1)
      dte2 = DateSerial(lngYr, 10, 31)
    Case 4
      'Between #1/1/00# And #12/31/00#
      dte1 = DateSerial(lngYr, 1, 1)
      dte2 = DateSerial(lngYr, 12, 31)
  End Select
  
  strSQL = "SELECT Year([dteDealDateIn]) AS SelectedYear, DatePart(""q"",[dteDealDateIn]) AS SelectedQ, " & _
    "Year([dteDealDateDisp]) AS YearClosed, Format([dteDealDateDisp],""q"") AS QClosed, tblDeal.* " & _
    "FROM tblDeal WHERE (((tblDeal.dteDealDateDisp) Between #" & dte1 & "# And #" & dte2 & "#) AND " & _
    "((tblDeal.lngDealDispNum)=1)) ORDER BY Year([dteDealDateIn]) DESC , DatePart(""q"",[dteDealDateIn]) " & _
    "DESC , tblDeal.dteDealDateDisp DESC;"

  strQry = "qry" & lngYr & "LTM-Q" & iQtr & "EndClosedDealData"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
  
End Function

Function SQLQr_LTMQEndClosedDealsPerMonth(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "LTM-Q" & iQtr & "EndClosedDealData" 'xqry2001LTM-Q4EndClosedDealData
 
  strSQL = "SELECT [" & strSubQry1 & "].SelectedYear, [" & strSubQry1 & "].YearClosed, " & _
  "Format([" & strSubQry1 & "]![dteDealDateDisp],""mmm"") AS SelectedMonth, " & _
  "Count([" & strSubQry1 & "].lngDealNum) AS ClosedDeals, True AS Closed, " & _
  "Month([" & strSubQry1 & "]![dteDealDateDisp]) AS MonthNum FROM [" & _
  strSubQry1 & "] GROUP BY [" & strSubQry1 & "].SelectedYear, " & _
  "[qry2004LTM-Q1EndClosedDealData].YearClosed, Format([" & strSubQry1 & _
  "]![dteDealDateDisp],""mmm""), True, Month([" & strSubQry1 & "]![dteDealDateDisp]) " & _
  "ORDER BY Month([" & strSubQry1 & "]![dteDealDateDisp]);"
 
  strQry = "qry" & lngYr & "LTM-Q" & iQtr & "EndClosedDealsPerMonth"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
 
End Function

Function SQLQr_LTMQEndCountOfSourceType(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "LTM-Q" & iQtr & "EndDealData" 'xqry2001LTM-Q4EndDealData

  strSQL = "SELECT DISTINCT Count([" & strSubQry1 & "].[lngDealNum]) AS [Deal Count], " & _
    "[tlkpSourceType].[txtSourceTypeDesc] FROM [" & strSubQry1 & "] INNER JOIN (tblSource INNER JOIN " & _
    "tlkpSourceType ON [tblSource].[lngSourceTypeNum]=[tlkpSourceType].[lngSourceTypeNum]) ON [" & _
    strSubQry1 & "].[lngDealSourceNum]=[tblSource].[lngSourceNum] " & _
    "GROUP BY [tlkpSourceType].[txtSourceTypeDesc] HAVING (((Count([" & strSubQry1 & "].lngDealNum)) > 0)) " & _
    "ORDER BY Count([" & strSubQry1 & "].[lngDealNum]) DESC;"
  
  strQry = "qry" & lngYr & "LTM-Q" & iQtr & "EndCountOfSourceType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLQr_LTMQEndDealData(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim dte1 As Date, dte2 As Date
  
  Select Case iQtr
    Case 1
      'e.g: Between #4/1/00# And #3/31/01#
      dte1 = DateSerial(lngYr - 1, 4, 1)
      dte2 = DateSerial(lngYr, 3, 31)
    Case 2
      'Between #7/1/00# And #6/30/01#
      dte1 = DateSerial(lngYr - 1, 7, 1)
      dte2 = DateSerial(lngYr, 6, 30)
    Case 3
      'Between #10/1/00# And #9/30/01#
      dte1 = DateSerial(lngYr - 1, 10, 1)
      dte2 = DateSerial(lngYr, 9, 30)
    Case 4
      'Between #1/1/00# And #12/31/00#
      dte1 = DateSerial(lngYr, 1, 1)
      dte2 = DateSerial(lngYr, 12, 31)
  End Select

  strSQL = "SELECT Year([dteDealDateIn]) AS SelectedYear, DatePart(""q"",[dteDealDateIn]) AS SelectedQ, " & _
    "tblDeal.* FROM tblDeal WHERE (((tblDeal.dteDealDateIn) Between #" & dte1 & "# And #" & dte2 & "#)) " & _
    "ORDER BY Year([dteDealDateIn]) DESC , DatePart(""q"",[dteDealDateIn]) DESC , tblDeal.dteDealDateIn DESC;"
  
  strQry = "qry" & lngYr & "LTM-Q" & iQtr & "EndDealData"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLQr_LTMQEndReviewedDealsPerMonth(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "LTM-Q" & iQtr & "EndDealData" 'xqry2001LTM-QxEndDealData

  strSQL = "SELECT [" & strSubQry1 & "].SelectedYear, Month([" & strSubQry1 & "]![dteDealDateIn]) " & _
    "AS MonthNum, Format([" & strSubQry1 & "]![dteDealDateIn],""mmm"") AS SelectedMonth, " & _
    "Count([" & strSubQry1 & "].lngDealNum) AS ReviewedDeals FROM [" & strSubQry1 & "] " & _
    "GROUP BY [" & strSubQry1 & "].SelectedYear, Month([" & strSubQry1 & "]![dteDealDateIn]), " & _
    "Format([" & strSubQry1 & "]![dteDealDateIn],""mmm"") " & _
    "ORDER BY [" & strSubQry1 & "].SelectedYear, Month([" & strSubQry1 & "]![dteDealDateIn]);"

  strQry = "qry" & lngYr & "LTM-Q" & iQtr & "EndReviewedDealsPerMonth"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLQr_LTMQEndShareOfSourceType(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String

  strSubQry1 = "qry" & lngYr & "LTM-Q" & iQtr & "EndCountOfSourceType"  'xqry2001LTM-Q4EndCountOfSourceType
  strSubQry2 = "qry" & lngYr & "LTM-Q" & iQtr & "TotReviewedDeals" 'xqry2001LTM-Q4TotReviewedDeals
  
  strSQL = "SELECT [" & strSubQry1 & "].[txtSourceTypeDesc] AS [Source Type], [" & strSubQry1 & _
    "]![Deal Count]/[" & strSubQry2 & "]![ReviewedDeals] AS Share " & _
    "FROM [" & strSubQry1 & "], [" & strSubQry2 & "] " & _
    "ORDER BY [" & strSubQry1 & "]![Deal Count]/[" & strSubQry2 & "]![ReviewedDeals] DESC;"
 
  strQry = "qry" & lngYr & "LTM-Q" & iQtr & "EndShareOfSourceType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  
End Function

Function SQLQr_LTMQTotClosedDeals(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "LTM-Q" & iQtr & "EndClosedDealData"    'xqry2001QxEndClosedDealData

  strSQL = "SELECT [" & strSubQry1 & "].SelectedYear, Count([" & strSubQry1 & "].lngDealNum) " & _
    "AS ClosedDeals FROM [" & strSubQry1 & "] GROUP BY [" & strSubQry1 & "].SelectedYear " & _
    "ORDER BY Count([" & strSubQry1 & "].lngDealNum);"

  strQry = "qry" & lngYr & "LTM-Q" & iQtr & "TotClosedDeals"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_LTMQEndTotReviewedDeals(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "LTM-Q" & iQtr & "EndDealData"    'xqry2001QxEndDealData

  strSQL = "SELECT Count([" & strSubQry1 & "].[lngDealNum]) AS ReviewedDeals " & _
           "FROM [" & strSubQry1 & "] ORDER BY Count([" & strSubQry1 & "].[lngDealNum]);"
  
  strQry = "qry" & lngYr & "LTM-Q" & iQtr & "TotReviewedDeals"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_AggregInvAndAverSize(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "DealData"            'xqry2001QxDealData
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "TotReviewedDeals"    'xQxTotReviewedDeals
  'cstNoWarrantsQ  = " & cstNoWarrantsQ  & "
                    
  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, " & _
    strSubQry2 & ".ReviewedDeals AS Transactions, Sum(IIf([" & cstNoWarrantsQ & "]![sglAmtOffered]=0, [" & _
    cstNoWarrantsQ & "]![sglTrancheSize],[" & cstNoWarrantsQ & "]![sglAmtOffered])) " & _
    "AS AggregateInvPotential, Sum(IIf([" & cstNoWarrantsQ & "]![sglAmtOffered]=0, [" & _
    cstNoWarrantsQ & "]![sglTrancheSize],[" & cstNoWarrantsQ & "]![sglAmtOffered]))/[" & _
    strSubQry2 & "]![ReviewedDeals] AS AverageSize " & _
    "FROM " & strSubQry2 & ", " & strSubQry1 & " INNER JOIN " & cstNoWarrantsQ & " ON " & _
    strSubQry1 & ".lngDealNum = " & cstNoWarrantsQ & ".lngSecDealNum " & _
    "GROUP BY " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, " & strSubQry2 & ".ReviewedDeals;"
  
  strQry = "qry" & lngYr & "Q" & iQtr & "AggregInvAndAverSize"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
  
End Function

Function SQLQr_ClosedDealData(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "ClosedDealData"

  strSQL = "SELECT " & strSubQry1 & ".* FROM " & _
           strSubQry1 & " WHERE ((" & strSubQry1 & ".QClosed=" & iQtr & "));"

  strQry = "qry" & lngYr & "Q" & iQtr & "ClosedDealData"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_DealData(lngYr As Long, iQ As Integer)
  Dim strSQL As String, strQry As String

  strSQL = "SELECT Year([dteDealDateIn]) AS SelectedYear, DatePart(""q"",[dteDealDateIn]) AS SelectedQ," & _
        " DatePart(""q"",[dteDealDateDisp]) AS QDisposed, tblDeal.* FROM tblDeal WHERE " & _
        "(((Year([dteDealDateIn])) = " & lngYr & ") And ((DatePart(""q"", [dteDealDateIn])) = " & iQ & _
        ")) ORDER BY tblDeal.dteDealDateIn DESC;"
        
  strQry = "qry" & lngYr & "Q" & iQ & "DealData"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_TotClosedDeals(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
    Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "ClosedDealData"              'x2001Q1ClosedDealData
  
  strSQL = "SELECT " & strSubQry1 & ".SelectedYear,  " & strSubQry1 & ".QClosed, " & _
    "Count(" & strSubQry1 & ".lngDealNum) AS ClosedDeals " & _
    "FROM " & strSubQry1 & " GROUP BY " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".QClosed;"
  
  strQry = "qry" & lngYr & "Q" & iQtr & "TotClosedDeals"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
  
End Function

Function SQLYr_ReviewedDealsPerMonth(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "DealData"             'xqry2001DealData

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, Format([" & strSubQry1 & "]![dteDealDateIn],'mmm') " & _
    "AS SelectedMonth, Count(" & strSubQry1 & ".lngDealNum) AS ReviewedDeals, False AS Closed, " & _
    "Month([" & strSubQry1 & "]![dteDealDateIn]) AS MonthNum " & _
    "FROM " & strSubQry1 & " " & _
    "GROUP BY " & strSubQry1 & ".SelectedYear, Format([" & strSubQry1 & "]![dteDealDateIn],'mmm'), False, " & _
    "Month([" & strSubQry1 & "]![dteDealDateIn]) " & _
    "ORDER BY Month([" & strSubQry1 & "]![dteDealDateIn]), Count(" & strSubQry1 & ".lngDealNum);"

  strQry = "qry" & lngYr & "ReviewedDealsPerMonth"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  
End Function

Function SQLQr_ShareOfCoverageLevel(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "CountOfIssuersWithFinData"
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "IssuerCoverage"
  
  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, " & strSubQry2 & _
    ".Group AS [Interest Coverage Range],[" & strSubQry2 & "]![Issuers]/[" & strSubQry1 & _
    "]![Issuers] AS [Share of Issuers Per Range] FROM " & strSubQry1 & ", " & strSubQry2 & ";"
  
  strQry = "qry" & lngYr & "Q" & iQtr & "ShareOfCoverageLevel"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)

End Function

Function SQLQr_ShareOfLeverageLevel(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "CountOfIssuersWithFinData"
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "IssuerLeverage"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, " & strSubQry2 & _
    ".Group AS [Debt Leverage Range],[" & strSubQry2 & "]![Issuers]/[" & strSubQry1 & _
    "]![Issuers] AS [Share of Issuers Per Range] FROM " & strSubQry1 & ", " & strSubQry2 & ";"

  strQry = "qry" & lngYr & "Q" & iQtr & "ShareOfLeverageLevel"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)

End Function

Function SQLYr_ShareOfDealType(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "CountOfDealType"
  strSubQry2 = "qry" & lngYr & "TotReviewedDeals"
  
  strSQL = "SELECT " & strSubQry1 & ".[Deal Type], [" & strSubQry1 & "]![Deal Type Count]/[" & strSubQry2 & _
    "]![ReviewedDeals] AS Share FROM " & strSubQry2 & ", " & strSubQry1 & _
    " ORDER BY [" & strSubQry1 & "]![Deal Type Count]/[" & strSubQry2 & "]![ReviewedDeals] DESC;"
 
  strQry = "qry" & lngYr & "ShareOfDealType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLQr_ShareOfDealType(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "CountOfDealType"
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "TotReviewedDeals"
  
  strSQL = "SELECT " & strSubQry1 & ".[Deal Type], [" & strSubQry1 & "]![Deal Type Count]/[" & strSubQry2 & _
    "]![ReviewedDeals] AS Share FROM " & strSubQry2 & ", " & strSubQry1 & _
    " ORDER BY [" & strSubQry1 & "]![Deal Type Count]/[" & strSubQry2 & "]![ReviewedDeals] DESC;"
 
  strQry = "qry" & lngYr & "Q" & iQtr & "ShareOfDealType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLYr_ShareOfLeverageLevel(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "CountOfIssuersWithFinData"
  strSubQry2 = "qry" & lngYr & "IssuerLeverage"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry2 & _
          ".Group AS [Debt Leverage Range],[" & strSubQry2 & "]![Issuers]/[" & strSubQry1 & _
       "]![Issuers Count] AS [Share of Issuers Per Range] FROM " & strSubQry1 & ", " & strSubQry2 & _
        " WHERE (" & strSubQry1 & ".SelectedYear = " & lngYr & ") ORDER BY " & strSubQry2 & ".Group;"
  
  
  strQry = "qry" & lngYr & "ShareOfLeverageLevel"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLYr_ShareOfCoverageLevel(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "CountOfIssuersWithFinData"
  strSubQry2 = "qry" & lngYr & "IssuerCoverage"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry2 & ".Group AS " & _
          "[Interest Coverage Range], [" & strSubQry2 & "]![Issuers]/[" & strSubQry1 & _
          "]![Issuers Count] AS [Share of Issuers Per Range] FROM " & strSubQry1 & ", " & strSubQry2 & _
          " WHERE (" & strSubQry1 & ".SelectedYear = " & lngYr & ") ORDER BY " & strSubQry2 & ".Group;"
  
  strQry = "qry" & lngYr & "ShareOfCoverageLevel"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLYr_ShareOfSecType(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "CountOfSecType"
  strSubQry2 = "qry" & lngYr & "TotReviewedDeals"
                                                      
  strSQL = "SELECT " & strSubQry1 & ".[Sec Type], [" & strSubQry1 & "]![Sec Type Count]/[" & _
    strSubQry2 & "]![ReviewedDeals] AS Share FROM " & strSubQry2 & ", " & strSubQry1 & _
    " ORDER BY [" & strSubQry1 & "]![Sec Type Count]/[" & strSubQry2 & "]![ReviewedDeals] DESC;"
    
  strQry = "qry" & lngYr & "ShareOfSecType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLQr_ShareOfSecType(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "CountOfSecType"
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "TotReviewedDeals"
                                                      
  strSQL = "SELECT [" & strSubQry1 & "].[Sec Type], [" & strSubQry1 & "]![Sec Type Count]/[" & _
    strSubQry2 & "]![ReviewedDeals] AS Share FROM " & strSubQry2 & ", " & strSubQry1 & _
    " ORDER BY [" & strSubQry1 & "]![Sec Type Count]/[" & strSubQry2 & "]![ReviewedDeals] DESC;"
    
  strQry = "qry" & lngYr & "Q" & iQtr & "ShareOfSecType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLYr_ShareOfSourceType(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "CountOfSourceType"
  strSubQry2 = "qry" & lngYr & "TotReviewedDeals"

  strSQL = "SELECT " & strSubQry1 & ".txtSourceTypeDesc AS [Source Type], [" & strSubQry1 & _
    "]![Deal Count]/[" & strSubQry2 & "]![ReviewedDeals] AS Share FROM " & strSubQry1 & ", " & strSubQry2 & _
    " ORDER BY [" & strSubQry1 & "]![Deal Count]/[" & strSubQry2 & "]![ReviewedDeals] DESC;"
  
  strQry = "qry" & lngYr & "ShareOfSourceType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  
End Function

Function SQLQr_ShareOfSourceType(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "CountOfSourceType"      'xqry2001Q1CountOfSourceType
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "TotReviewedDeals"     'xqry2001Q1TotReviewedDeals
  
  strSQL = "SELECT [" & strSubQry1 & "].[txtSourceTypeDesc] AS [Source Type], [" & strSubQry1 & _
    "]![Deal Count]/[" & strSubQry2 & "]![ReviewedDeals] AS Share FROM " & strSubQry1 & ", " & strSubQry2 & _
    " ORDER BY [" & strSubQry1 & "]![Deal Count]/[" & strSubQry2 & "]![ReviewedDeals] DESC;"
  
  strQry = "qry" & lngYr & "Q" & iQtr & "ShareOfSourceType"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
  
End Function

Function SQLQr_SourcesWith1Deal(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String, strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "DealData"             'xqry200xQxDealData
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "SourcesWith1DealIDs"  'xqry200xQxSourcesWith1DealIDs

  strSQL = "SELECT DISTINCT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, "
  strSQL = strSQL & "tblSource.txtSourceName AS Source, Count(CLng([lngDealSourceNum])) AS [Deal Count], "
  strSQL = strSQL & strSubQry2 & ".SrceID FROM ( " & strSubQry1 & " INNER JOIN tblSource ON "
  strSQL = strSQL & strSubQry1 & ".lngDealSourceNum = tblSource.lngSourceNum) INNER JOIN "
  strSQL = strSQL & strSubQry2 & " ON tblSource.txtSourceName = " & strSubQry2 & ".Source GROUP BY "
  strSQL = strSQL & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, tblSource.txtSourceName, "
  strSQL = strSQL & strSubQry2 & ".SrceID HAVING (((Count(CLng([lngDealSourceNum]))) = 1))"
  strSQL = strSQL & "ORDER BY tblSource.txtSourceName;"

' strSQL = "SELECT DISTINCT " & strSubQry1 & ".SelectedYear, " & strSubQry2 & ".Source, " & _
'    "Count(CLng([lngDealSourceNum])) AS [Deal Count] FROM " & strSubQry1 & " INNER JOIN " & _
'    strSubQry2 & " ON " & strSubQry1 & ".lngDealSourceNum = " & strSubQry2 & ".SrceID GROUP BY " & _
'    strSubQry1 & ".SelectedYear, " & strSubQry2 & ".Source " & _
'    "HAVING (Count(CLng([lngDealSourceNum]))=1) ORDER BY qry2003SourcesWith1DealIDs.Source;"

  strQry = "qry" & lngYr & "Q" & iQtr & "SourcesWith1Deal"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
  
End Function

Function SQLQr_SourcesWith1DealSub(lngYr As Long, iQ As Integer)
  Dim strSQL As String, strQry As String, strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQ & "DealData"
  strSubQry2 = cstSourcesWithDeals
  
  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, " & strSubQry2 & _
          ".txtSourceName AS Source, Count(" & strSubQry2 & ".txtSourceName) AS [Deal Count], " & _
          "CLng([lngDealSourceNum]) AS SrceID FROM " & strSubQry2 & " INNER JOIN " & strSubQry1 & _
          " ON " & strSubQry2 & ".lngSourceNum = " & strSubQry1 & ".lngDealSourceNum GROUP BY " & _
          strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, " & strSubQry2 & _
          ".txtSourceName, CLng([lngDealSourceNum]) HAVING (((Count(" & _
          strSubQry2 & ".txtSourceName))=1)) ORDER BY " & strSubQry2 & ".txtSourceName;"
  
  strQry = "qry" & lngYr & "Q" & iQ & "SourcesWith1DealIDs"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)

End Function

Function SQLQr_SourcesWith2PlusDeals(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String, strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "DealData"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, " & _
           "tblsource.txtSourceName AS Source, Count(tblsource.txtSourceName) AS [Deal Count] " & _
           "FROM " & strSubQry1 & " LEFT JOIN tblSource ON " & strSubQry1 & ".lngDealSourceNum = " & _
           "tblsource.lngSourceNum GROUP BY " & strSubQry1 & ".SelectedYear, " & strSubQry1 & _
           ".SelectedQ, tblsource.txtSourceName HAVING ((Count(" & "tblsource.txtSourceName))>=2) " & _
           "ORDER BY Count(tblsource.txtSourceName) DESC;"

  strQry = "qry" & lngYr & "Q" & iQtr & "SourcesWith2PlusDeals"
  'Debug.Print strSQL
  Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_TotClosedDealsForMultiDealSources(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "ClosedDealData"
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "SourcesWith2PlusDeals"

  strSQL = "SELECT " & strSubQry1 & ".YearClosed, " & strSubQry2 & ".Source, " & strSubQry1 & ".QClosed, "
  strSQL = strSQL & strSubQry1 & ".SelectedYear, Count(" & strSubQry1 & ".lngDealNum) AS ClosedDeals FROM "
  strSQL = strSQL & strSubQry2 & " LEFT JOIN (tblSource LEFT JOIN " & strSubQry1 & " ON tblSource.lngSourceNum = "
  strSQL = strSQL & strSubQry1 & ".lngDealSourceNum) ON " & strSubQry2
  strSQL = strSQL & ".Source = tblSource.txtSourceName GROUP BY " & strSubQry1 & ".YearClosed," & strSubQry2
  strSQL = strSQL & ".Source, " & strSubQry1 & ".QClosed, " & strSubQry1 & ".SelectedYear;"

  strQry = "qry" & lngYr & "Q" & iQtr & "TotClosedDealsForMultiDealSources"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_TotClosedDealsForSingleDealSources(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "SourcesWith1Deal"       'xqry2001SourcesWith1Deal
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "ClosedDealData"         'xqry2001DealData
   
  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, 'Sources With One Deal'"
  strSQL = strSQL & "AS Source, Count(" & strSubQry2 & ".lngDealNum) AS ClosedDeals FROM " & strSubQry1
  strSQL = strSQL & " LEFT JOIN " & strSubQry2 & " ON " & strSubQry1 & ".SrceID = " & strSubQry2
  strSQL = strSQL & ".lngDealSourceNum GROUP BY " & strSubQry1 & ".SelectedYear, " & strSubQry1
  strSQL = strSQL & ".SelectedQ, 'Sources With One Deal';"
  
  strQry = "qry" & lngYr & "Q" & iQtr & "TotClosedDealsForSingleDealSources"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)

End Function

Function SQLQr_TotSourcesWith1Deal(lngYr As Long, iQtr As Integer)
 Dim strSQL As String, strQry As String, strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "SourcesWith1DealIDs"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, sum(" & strSubQry1 & ".[Deal Count]) AS " & _
           "[Sources with 1 deal] FROM " & strSubQry1 & " GROUP BY " & strSubQry1 & ".SelectedYear;"
  
  strQry = "qry" & lngYr & "Q" & iQtr & "TotSourcesWith1Deal"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_TotSourcesWith2PlusDeals(lngYr As Long, iQtr As Integer)
 Dim strSQL As String, strQry As String, strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "SourcesWith2PlusDeals"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, "
  strSQL = strSQL & "Count(" & strSubQry1 & ".[Deal Count]) AS [MultiDeal Sources], "
  strSQL = strSQL & "Sum(" & strSubQry1 & ".[Deal Count]) AS [Reviewed Deals] FROM " & strSubQry1
  strSQL = strSQL & " GROUP BY " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ;"
  
  strQry = "qry" & lngYr & "Q" & iQtr & "TotSourcesWith2PlusDeals"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_SourcesWith1DealAverSize(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String, strSubQry1 As String, strSubQry2 As String, strSubQry3 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "SourcesWith1Deal"      'qryYYYYQxSourcesWith1Deal
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "TotSourcesWith1Deal"   'qryYYYYQxTotSourcesWith1Deal
  strSubQry3 = "qry" & lngYr & "Q" & iQtr & "DealData"              'qryYYYYQxDealData

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, 'Single Deal Sources' AS "
  strSQL = strSQL & "Sources, Sum([qryNonWarrantSecuritiesData]![sglAmtOffered])/[" & strSubQry2
  strSQL = strSQL & "]![Sources with 1 deal] AS DealAverSize, " & strSubQry2
  strSQL = strSQL & ".[Sources with 1 deal] AS [Reviewed Deals] FROM (" & strSubQry2
  strSQL = strSQL & " INNER JOIN " & strSubQry1 & " ON " & strSubQry2 & ".SelectedYear =" & strSubQry1
  strSQL = strSQL & ".SelectedYear) INNER JOIN (" & strSubQry3 & " LEFT JOIN qryNonWarrantSecuritiesData ON "
  strSQL = strSQL & strSubQry3 & ".lngDealNum = qryNonWarrantSecuritiesData.lngSecDealNum) ON "
  strSQL = strSQL & strSubQry1 & ".SrceID = " & strSubQry3 & ".lngDealSourceNum GROUP BY "
  strSQL = strSQL & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, 'Single Deal Sources',"
  strSQL = strSQL & strSubQry2 & ".[Sources with 1 deal];"

  'Debug.Print strSQL
  strQry = "qry" & lngYr & "Q" & iQtr & "SourcesWith1DealAverSize"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_SourcesWith2PlusDealsSecData(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String, strSubQry1 As String, strSubQry2 As String
 
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "SourcesWith2PlusDeals"
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "DealData"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, " & strSubQry1
  strSQL = strSQL & ".Source, " & strSubQry1 & ".[Deal Count] AS [Reviewed Deals], "
  strSQL = strSQL & "qryNonWarrantSecuritiesData.sglAmtOffered FROM "
  strSQL = strSQL & strSubQry1 & " INNER JOIN ((tblSource LEFT JOIN " & strSubQry2
  strSQL = strSQL & " ON tblSource.lngSourceNum =  " & strSubQry2 & ".lngDealSourceNum) LEFT JOIN "
  strSQL = strSQL & "qryNonWarrantSecuritiesData ON  " & strSubQry2 & ".lngDealNum = "
  strSQL = strSQL & "qryNonWarrantSecuritiesData.lngSecDealNum) ON " & strSubQry1
  strSQL = strSQL & ".Source = tblSource.txtSourceName GROUP BY "
  strSQL = strSQL & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedQ, " & strSubQry1
  strSQL = strSQL & ".Source, " & strSubQry1 & ".[Deal Count], qryNonWarrantSecuritiesData.sglAmtOffered;"

  strQry = "qry" & lngYr & "Q" & iQtr & "SourcesWith2PlusDealsSecData"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
  
End Function

Function SQLYr_TotClosedDeals(lngYr As Long) ' 90
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "ClosedDealData"            'xqry2001ClosedDealData

  strSQL = "SELECT  " & strSubQry1 & ".SelectedYear, Count( " & strSubQry1 & ".lngDealNum) AS ClosedDeals " & _
    "FROM  " & strSubQry1 & " GROUP BY  " & strSubQry1 & ".SelectedYear " & _
    "ORDER BY Count( " & strSubQry1 & ".lngDealNum);"

  strQry = "qry" & lngYr & "TotClosedDeals"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLYr_TotReviewedDeals(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "DealData"

  strSQL = "SELECT Count(" & strSubQry1 & ".lngDealNum) AS ReviewedDeals FROM " & strSubQry1 & _
           " ORDER BY Count(" & strSubQry1 & ".lngDealNum);"

  strQry = "qry" & lngYr & "TotReviewedDeals"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  
End Function

Function SQLYr_TotSourcesWith1Deal(lngYr As Long)
 Dim strSQL As String, strQry As String, strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "TotSourcesWith1DealSub"
  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, Sum(" & strSubQry1 & ".SingleDealSources) AS " & _
           "[Sources with 1 deal] FROM " & strSubQry1 & " GROUP BY " & strSubQry1 & ".SelectedYear;"

  strQry = "qry" & lngYr & "TotSourcesWith1Deal"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLYr_TotSourcesWith1DealSub(lngYr As Long)
 Dim strSQL As String, strQry As String, strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "SourcesWith1DealIDs"
  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, Count(" & strSubQry1 & ".Source) AS SingleDealSources " & _
           "FROM " & strSubQry1 & " GROUP BY " & strSubQry1 & ".SelectedYear;"

  strQry = "qry" & lngYr & "TotSourcesWith1DealSub"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLYr_UNClosedVsReviewedDealsPerMonth(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "ReviewedDealsPerMonth"
  strSubQry2 = "qry" & lngYr & "ClosedDealsPerMonth"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".SelectedMonth, " & _
          strSubQry1 & ".ReviewedDeals, " & _
          "IIf(IsNull(" & strSubQry1 & ".SelectedMonth),0," & strSubQry2 & ".ClosedDeals) " & _
          "AS ClosedDeals FROM " & strSubQry1 & " LEFT JOIN " & strSubQry2 & " ON (" & _
          strSubQry1 & ".SelectedYear = " & strSubQry2 & ".YearClosed) AND (" & _
          strSubQry1 & ".SelectedMonth = " & strSubQry2 & ".SelectedMonth)" & _
          " ORDER BY " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".MonthNum;"
      
  strQry = "qry" & lngYr & "UNClosedVsReviewedDealsPerMonth"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLYr_SourcesWithDealID(lngYr As Long)
  Dim strSQL As String, strQry As String

  strSQL = "SELECT DISTINCT tblSource.*, tblDeal.*, Year([dteDealDateIn]) AS SelectedYear, " & _
          "DatePart(""q"",[dteDealDateIn]) AS SelectedQ, Year([dteDealDateDisp]) AS YearClosed, " & _
          "DatePart(""q"",[dteDealDateDisp]) AS QDisposed  " & _
          "FROM tblSource LEFT JOIN tblDeal ON tblSource.lngSourceNum = tblDeal.lngDealSourceNum " & _
          "WHERE (((Year([dteDealDateIn]))=" & lngYr & ") AND ((tblDeal.lngDealNum) Is Not Null)) " & _
          "OR (((Year([dteDealDateDisp]))=" & lngYr & ") AND ((tblDeal.lngDealDispNum)=1));"

  strQry = "qry" & lngYr & "SourcesWithDealID"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLQr_LTMQEndClosedVsReviewedDeals(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "LTM-Q" & iQtr & "EndReviewedDealsPerMonth"
  strSubQry2 = "qry" & lngYr & "LTM-Q" & iQtr & "EndClosedDealsPerMonth"

  strSQL = "SELECT [" & strSubQry1 & "].SelectedYear, [" & strSubQry1 & "].SelectedMonth, [" & strSubQry1 & _
    "].ReviewedDeals, IIf(IsNull([" & strSubQry2 & "]![SelectedMonth]),0,[" & strSubQry2 & "]!" & _
    "[ClosedDeals]) AS ClosedDeals FROM [" & strSubQry1 & "] LEFT JOIN [" & strSubQry2 & "] ON ([" & strSubQry1 & _
    "].SelectedYear = [" & strSubQry2 & "].YearClosed) AND ([" & strSubQry1 & "].SelectedMonth = [" & _
    strSubQry2 & "].SelectedMonth) ORDER BY [" & strSubQry1 & "].SelectedYear, [" & strSubQry1 & "].MonthNum;"

  strQry = "qry" & lngYr & "LTM-Q" & iQtr & "EndClosedVsReviewedDeals"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)

End Function

Function SQLYr_UNPrevYearDealSizeComp(lngYr As Long)
  Dim strSQL As String, strQry As String, strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "AggregInvAndAverSize"
  strSubQry2 = "qry" & (lngYr - 1) & "AggregInvAndAverSize"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear," & strSubQry1 & ".Transactions, " & _
    strSubQry1 & ".AggregateInvPotential, " & strSubQry1 & ".AverageSize FROM " & strSubQry1 & " " & _
    "UNION SELECT " & strSubQry2 & ".SelectedYear," & strSubQry2 & ".Transactions, " & _
    strSubQry2 & ".AggregateInvPotential, " & strSubQry2 & ".AverageSize FROM " & strSubQry2 & " " & _
    "ORDER BY " & strSubQry1 & ".SelectedYear DESC;"
  
  strQry = "qry" & lngYr & "UNPrevYearDealSizeComp"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLQr_UNPrevYearQEndDealSizeComp(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String, strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "AggregInvAndAverSize"
  strSubQry2 = "qry" & (lngYr - 1) & "Q" & iQtr & "AggregInvAndAverSize"
  
  strSQL = "SELECT  " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".Transactions, " & _
      strSubQry1 & ".AggregateInvPotential,  " & strSubQry1 & ".AverageSize " & _
      "FROM  " & strSubQry1 & " " & _
      "UNION SELECT  " & strSubQry2 & ".SelectedYear, " & strSubQry2 & ".Transactions, " & _
      strSubQry2 & ".AggregateInvPotential,  " & strSubQry2 & ".AverageSize " & _
      "FROM  " & strSubQry2 & " " & _
      "ORDER BY  " & strSubQry1 & ".SelectedYear;"
  
  strQry = "qry" & lngYr & "UNPrevYearQ" & iQtr & "DealSizeComp"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLYr_UNSourcesSummary(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String

 strSubQry1 = "qry" & lngYr & "SummaryMultiDealSources"
 strSubQry2 = "qry" & lngYr & "SummarySingleDealSources"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".Source, " & strSubQry1 & _
          ".[Reviewed Deals], " & strSubQry1 & ".ClosedDeals, " & strSubQry1 & ".DealAverSize, " & _
          strSubQry1 & ".ListOrder FROM " & strSubQry1 & " UNION SELECT " & strSubQry2 & _
          ".SelectedYear, " & strSubQry2 & ".Source, " & strSubQry2 & ".[Reviewed Deals], " & _
          strSubQry2 & ".ClosedDeals, " & strSubQry2 & ".DealAverSize, " & strSubQry2 & _
          ".ListOrder FROM " & strSubQry2 & " ORDER BY " & strSubQry1 & ".ListOrder, " & strSubQry1 & _
          ".[Reviewed Deals] DESC , " & strSubQry1 & ".DealAverSize DESC , " & strSubQry1 & ".Source;"

  strQry = "qry" & lngYr & "UNSourcesSummary"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLYr_SourcesWith2PlusDeals(lngYr As Long)
  Dim strSQL As String, strQry As String, strSubQry1 As String, strSubQry2 As String

  strSubQry1 = "qry" & lngYr & "DealData"
  strSubQry2 = "qry" & lngYr & "SourcesWithDealID"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry2 & ".txtSourceName AS Source, " & _
    "Count(" & strSubQry2 & ".txtSourceName) AS [Deal Count] FROM " & strSubQry1 & _
    " INNER JOIN " & strSubQry2 & " ON (" & strSubQry1 & ".lngDealSourceNum = " & strSubQry2 & _
    ".lngSourceNum) AND (" & strSubQry1 & ".lngDealNum = " & strSubQry2 & ".lngDealNum) " & _
    "GROUP BY " & strSubQry1 & ".SelectedYear, " & strSubQry2 & ".txtSourceName " & _
    "HAVING ((Count(" & strSubQry2 & ".txtSourceName)) >= 2) ORDER BY " & _
    "Count(" & strSubQry2 & ".txtSourceName) DESC , qry2003SourcesWithDealID.txtSourceName;"

  strQry = "qry" & lngYr & "SourcesWith2PlusDeals"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLYr_SourcesWith1DealSub(lngYr As Long)
  Dim strSQL As String, strQry As String, strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "DealData"
  strSubQry2 = "qry" & lngYr & "SourcesWithDealID"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry2 & ".txtSourceName AS Source, " & _
    strSubQry2 & ".txtSourceContactLast, Count(" & strSubQry2 & ".txtSourceName) AS [Deal Count], " & _
    "CLng([lngDealSourceNum]) AS SrceID FROM " & strSubQry2 & " INNER JOIN " & strSubQry1 & " ON " & _
    strSubQry2 & ".lngSourceNum=" & strSubQry1 & ".lngDealSourceNum GROUP BY " & strSubQry1 & _
    ".SelectedYear, " & strSubQry2 & ".txtSourceName, " & strSubQry2 & ".txtSourceContactLast, " & _
    "CLng([lngDealSourceNum]) HAVING (Count(" & strSubQry2 & ".txtSourceName)=1) " & _
    "ORDER BY " & strSubQry2 & ".txtSourceName, " & strSubQry2 & ".txtSourceContactLast;"
  
  strQry = "qry" & lngYr & "SourcesWith1DealIDs"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)

End Function

Function SQLYr_SourcesWith1DealAverSize(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String, strSubQry3 As String
  
  strSubQry1 = "qry" & lngYr & "SourcesWith1DealIDs"
  strSubQry2 = "qry" & lngYr & "TotSourcesWith1Deal"
  strSubQry3 = "qry" & lngYr & "DealData"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, 'Single Deal Sources' AS Sources, " & _
    "Sum([" & cstNoWarrantsQ & "]![sglAmtOffered])/[" & strSubQry2 & "]!" & "[Sources with 1 deal] AS " & _
    "DealAverSize, " & strSubQry2 & ".[Sources with 1 deal] AS [Reviewed Deals] FROM " & _
    strSubQry2 & " INNER JOIN ( " & strSubQry1 & " INNER JOIN (" & strSubQry3 & _
    " LEFT JOIN " & cstNoWarrantsQ & " ON " & strSubQry3 & ".lngDealNum = " & _
    cstNoWarrantsQ & ".lngSecDealNum) ON " & strSubQry1 & ".SrceID = " & strSubQry3 & _
    ".lngDealSourceNum) ON " & strSubQry2 & ".SelectedYear = " & strSubQry1 & ".SelectedYear GROUP BY " & _
    strSubQry1 & ".SelectedYear, 'Single Deal Sources', " & strSubQry2 & ".[Sources with 1 deal];"
  
  strQry = "qry" & lngYr & "SourcesWith1DealAverSize"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLYr_SourcesWith1Deal(lngYr As Long)
  Dim strSQL As String, strQry As String, strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "SourcesWith1DealIDs"

  strSQL = "SELECT DISTINCT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".Source, " & _
    strSubQry1 & ".txtSourceContactLast, Count(" & strSubQry1 & ".SrceID) AS [Deal Count] " & _
    "FROM " & strSubQry1 & " GROUP BY " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".Source, " & _
    strSubQry1 & ".txtSourceContactLast ORDER BY " & strSubQry1 & ".Source;"
    
  strQry = "qry" & lngYr & "SourcesWith1Deal"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLYr_SourcesWith2PlusDealsSecData(lngYr As Long)
  Dim strSQL As String, strQry As String, strSubQry1 As String
  
  strSubQry1 = "qry" & lngYr & "SourcesWithDealID"

  strSQL = "SELECT " & strSubQry1 & ".YearClosed AS SelectedYear, " & strSubQry1 & _
    ".txtSourceName AS Source, Count(" & strSubQry1 & ".txtSourceName) AS [Deal Count], " & _
    strSubQry1 & ".lngSourceNum AS SrceID FROM " & strSubQry1 & " GROUP BY " & _
    strSubQry1 & ".YearClosed, " & strSubQry1 & ".txtSourceName, " & strSubQry1 & ".lngSourceNum " & _
    "HAVING (((Count(" & strSubQry1 & ".txtSourceName)) >= 2)) ORDER BY " & strSubQry1 & _
    ".txtSourceName, Count(" & strSubQry1 & ".txtSourceName) DESC;"

  strQry = "qry" & lngYr & "SourcesWith2PlusDealsSecData"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
  
End Function

Function SQLYr_TotClosedDealsForMultiDealSources(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "ClosedDealData"
  strSubQry2 = "qry" & lngYr & "SourcesWith2PlusDeals"

  strSQL = "SELECT " & strSubQry1 & ".YearClosed, " & strSubQry2 & ".Source, " & strSubQry1 & _
    ".SelectedYear, Count(" & strSubQry1 & ".lngDealNum) AS ClosedDeals FROM (" & strSubQry2 & _
    " LEFT JOIN tblSource ON " & strSubQry2 & ".Source = tblSource.txtSourceName) LEFT JOIN " & _
    strSubQry1 & " ON tblSource.lngSourceNum = " & strSubQry1 & ".lngDealSourceNum " & _
    "GROUP BY " & strSubQry1 & ".YearClosed," & strSubQry2 & ".Source, " & strSubQry1 & ".SelectedYear;"

  strQry = "qry" & lngYr & "TotClosedDealsForMultiDealSources"
   Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLYr_TotClosedDealsForSingleDealSources(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String
  
  strSubQry1 = "qry" & lngYr & "SourcesWith1DealIDs"
  strSubQry2 = "qry" & lngYr & "ClosedDealData"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".YearClosed, " & _
    "'Sources With One Deal' AS Source, " & " Count(" & strSubQry2 & ".lngDealNum) " & _
    "AS ClosedDeals FROM " & strSubQry1 & " INNER JOIN " & strSubQry2 & " ON " & strSubQry1 & _
    ".Source = " & strSubQry2 & ".Source GROUP BY " & strSubQry1 & ".SelectedYear, " & _
    strSubQry1 & ".YearClosed, 'Sources With One Deal';"
  
  strQry = "qry" & lngYr & "TotClosedDealsForSingleDealSources"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)

End Function

Function SQLYr_SummaryMultiDealSources(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String

  strSubQry1 = "qry" & lngYr & "SourcesWith2PlusDealsSecData"
  strSubQry2 = "qry" & lngYr & "TotClosedDealsForMultiDealSources"

  strSQL = "SELECT DISTINCT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".Source, " & _
            strSubQry1 & ".[Reviewed Deals], " & strSubQry2 & ".ClosedDeals, Sum([" & strSubQry1 & _
            "]![sglAmtOffered])/[" & strSubQry1 & "]![Reviewed Deals] AS DealAverSize, 1 AS ListOrder " & _
            "FROM " & strSubQry1 & " LEFT JOIN " & strSubQry2 & " ON " & strSubQry1 & ".Source = " & _
            strSubQry2 & ".Source GROUP BY " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".Source, " & _
            strSubQry1 & ".[Reviewed Deals], " & strSubQry2 & ".ClosedDeals, 1 ORDER BY " & strSubQry1 & _
            ".[Reviewed Deals] DESC , Sum([" & strSubQry1 & "]![sglAmtOffered])/[" & strSubQry1 & _
            "]![Reviewed Deals] DESC;"

  strQry = "qry" & lngYr & "SummaryMultiDealSources"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLYr_SummarySingleDealSources(lngYr As Long)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String

  strSubQry1 = "qry" & lngYr & "SourcesWith1DealAverSize"
  strSubQry2 = "qry" & lngYr & "TotClosedDealsForSingleDealSources"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry2 & ".Source, " & strSubQry1 & _
    ".[Reviewed Deals]," & strSubQry2 & ".ClosedDeals, " & strSubQry1 & ".DealAverSize, 2 AS ListOrder " & _
    "FROM " & strSubQry1 & ", " & strSubQry2 & " GROUP BY " & strSubQry1 & ".SelectedYear, " & _
    strSubQry2 & ".Source," & strSubQry1 & ".[Reviewed Deals], " & strSubQry2 & ".ClosedDeals, " & _
    strSubQry1 & ".DealAverSize, 2;"
  
  strQry = "qry" & lngYr & "SummarySingleDealSources"
  Call CreateDBQuery(strQry, strSQL, cstQryDescY, blnOverwrite)
End Function

Function SQLQr_UNSourcesSummary(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String

  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "SummaryMultiDealSources"
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "SummarySingleDealSources"

  strSQL = "SELECT " & strSubQry1 & ".SelectedYear," & iQtr & " as SelectedQ," & strSubQry1 & ".Source, "
  strSQL = strSQL & strSubQry1 & ".[Reviewed Deals]," & strSubQry1 & ".ClosedDeals," & strSubQry1
  strSQL = strSQL & ".DealAverSize," & strSubQry1 & ".ListOrder FROM " & strSubQry1 & " UNION SELECT "
  strSQL = strSQL & strSubQry2 & ".SelectedYear," & iQtr & ", " & strSubQry2 & ".Source," & strSubQry2
  strSQL = strSQL & ".[Reviewed Deals]," & strSubQry2 & ".ClosedDeals," & strSubQry2 & ".DealAverSize,"
  strSQL = strSQL & strSubQry2 & ".ListOrder FROM " & strSubQry2 & " ORDER BY " & strSubQry1 & ".ListOrder,"
  strSQL = strSQL & strSubQry1 & ".[Reviewed Deals] DESC ," & strSubQry1
  strSQL = strSQL & ".DealAverSize DESC ," & strSubQry1 & ".Source;"

  strQry = "qry" & lngYr & "Q" & iQtr & "UNSourcesSummary"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_SummaryMultiDealSources(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String

  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "SourcesWith2PlusDealsSecData"
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "TotClosedDealsForMultiDealSources"

  strSQL = "SELECT DISTINCT " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".Source, " & _
            strSubQry1 & ".[Reviewed Deals], " & strSubQry2 & ".ClosedDeals, Sum([" & strSubQry1 & _
            "]![sglAmtOffered])/[" & strSubQry1 & "]![Reviewed Deals] AS DealAverSize, 1 AS ListOrder " & _
            "FROM " & strSubQry1 & " LEFT JOIN " & strSubQry2 & " ON (" & strSubQry1 & ".SelectedYear = " & _
            strSubQry2 & ".YearClosed) AND (" & strSubQry1 & ".Source = " & strSubQry2 & ".Source) " & _
            "GROUP BY " & strSubQry1 & ".SelectedYear, " & strSubQry1 & ".Source, " & strSubQry1 & _
            ".[Reviewed Deals], " & strSubQry2 & ".ClosedDeals, 1 ORDER BY " & strSubQry1 & _
            ".[Reviewed Deals] DESC , Sum([" & strSubQry1 & "]![sglAmtOffered])/[" & strSubQry1 & _
            "]![Reviewed Deals] DESC;"
            
  strQry = "qry" & lngYr & "Q" & iQtr & "SummaryMultiDealSources"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Function SQLQr_SummarySingleDealSources(lngYr As Long, iQtr As Integer)
  Dim strSQL As String, strQry As String
  Dim strSubQry1 As String, strSubQry2 As String

  strSubQry1 = "qry" & lngYr & "Q" & iQtr & "SourcesWith1DealAverSize"
  strSubQry2 = "qry" & lngYr & "Q" & iQtr & "TotClosedDealsForSingleDealSources"
                                                
  strSQL = "SELECT " & strSubQry1 & ".SelectedYear, " & strSubQry2 & ".Source, " & strSubQry1
  strSQL = strSQL & ".[Reviewed Deals]," & strSubQry2 & ".ClosedDeals, " & strSubQry1
  strSQL = strSQL & ".DealAverSize, 2 AS ListOrder FROM " & strSubQry1 & " INNER JOIN "
  strSQL = strSQL & strSubQry2 & " ON " & strSubQry1 & ".SelectedYear = " & strSubQry2
  strSQL = strSQL & ".SelectedYear GROUP BY " & strSubQry1 & ".SelectedYear, " & strSubQry2 & ".Source,"
  strSQL = strSQL & strSubQry1 & ".[Reviewed Deals], " & strSubQry2 & ".ClosedDeals, " & strSubQry1
  strSQL = strSQL & ".DealAverSize, 2;"
                                                
  strQry = "qry" & lngYr & "Q" & iQtr & "SummarySingleDealSources"
   Call CreateDBQuery(strQry, strSQL, cstQryDescQ, blnOverwrite)
End Function

Public Sub SQLYr_NonWarrantSecuritiesData()
  Dim strSQL As String

  strSQL = "SELECT tblSecurity.* FROM tblSecurity WHERE (((tblSecurity.lngSecTypeNum)<>5));"
  Call CreateDBQuery(cstNoWarrantsQ, strSQL, cstDescAll, blnOverwrite)
  
End Sub

'-----------------------------------------------------------------
Public Function CreateDBQuery(strQryName As String, strSQL As String, strDesc As String, blnReplace As Boolean)
  Dim qdf As DAO.QueryDef
  Dim prop As DAO.Property
  
  On Error Resume Next
CreateNew:
  Set qdf = CurrentDb.CreateQueryDef(strQryName, strSQL) 'dbs set by first calling proc
  If Err = 0 Then
    Debug.Print "Qry: " & strQryName & " created"
  Else
    If Err = 3012 Then  'obj already exists
      Err.Clear
      If blnReplace Then
        CurrentDb.QueryDefs.Delete strQryName
        GoTo CreateNew
      Else
        Set qdf = CurrentDb.QueryDefs(strQryName)
      End If
    ElseIf Err = 3265 Then 'INTO tbl not found (MTBL qries): ignore
      Err.Clear
      Debug.Print "Err = 3265"
      Exit Function
    Else
      GoTo CreateDBQueryErr
    End If
  End If
  
  If Len(strDesc) > 0 Then
    Set prop = qdf.Properties("Description")
    If Err <> 0 Then
      If Err = 3270 Then 'Desc:  Property not found.   ' Or Err = 91
        Err.Clear
        On Error GoTo CreateDBQueryErr
        Set prop = qdf.CreateProperty("Description", dbText, strDesc)
        qdf.Properties.Append prop
      Else
        GoTo CreateDBQueryErr
      End If
    Else
      prop.Value = strDesc
    End If
  End If
  
CreateDBQueryExit:
  CurrentDb.QueryDefs.Refresh
  Set prop = Nothing
  Set qdf = Nothing
  Exit Function

CreateDBQueryErr:
  MsgBox strQryName & " Error: " & Err & vbCrLf & "Desc:  " & Err.Description, , "CreateDBQuery"
  Debug.Print strQryName & " Error: " & Err & vbCrLf & "Desc:  " & Err.Description
  Resume CreateDBQueryExit
End Function

Function DeleteDBQry(strQryName As String)
  On Error GoTo DeleteDBQryErr

  CurrentDb.QueryDefs.Delete strQryName
  
DeleteDBQryExit:
  Application.CurrentDb.QueryDefs.Refresh
  Exit Function

DeleteDBQryErr:
  If Err <> 3265 Then
    MsgBox strQryName & " Error: " & Err & vbCrLf & "Desc:  " & Err.Description, , "DeleteDBQry"
    Debug.Print strQryName & " Error: " & Err & vbCrLf & "Desc:  " & Err.Description
  Else
    Err.Clear
  End If
  Resume DeleteDBQryExit
End Function

Function RunDBQuery(strQryName As String)
  On Error GoTo RunDBQueryErr
  CurrentDb.Execute strQryName
  
RunDBQueryExit:
  Application.CurrentDb.QueryDefs.Refresh
  Exit Function
    
RunDBQueryErr:
  If Err = 3010 Then
   Err.Clear
  Else
    MsgBox "Error: " & Err.Number & " : " & Err.Description, vbExclamation, "RunDBQuery"
  End If
  Resume RunDBQueryExit
End Function
