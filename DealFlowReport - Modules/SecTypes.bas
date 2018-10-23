Attribute VB_Name = "SecTypes"
Option Compare Database
Option Explicit
'
'*****************************************************************
' DF Reports-SecTypes Sep-24-02 10:35
'******************************************************************
'

Public Function GetSecTypesStr(varDeal As Variant) As String
  Dim dbs As DAO.Database
  Dim rst As DAO.Recordset
  Dim strSQL As String, str As String, strType As String, strPrevType As String
  Dim lngDeal As Long
  Dim i As Integer, iTot As Integer
  
  str = "": strSQL = "": strType = "": strPrevType = "": i = 0: iTot = 0:
  On Error GoTo GetSecTypesStrErr
  
  If IsNull(varDeal) Then GoTo GetSecTypesStrExit
  lngDeal = CLng(varDeal)
  
  strSQL = "SELECT tblSecurity.lngSecDealNum AS Deal, tlkpSecType.txtSecType " & _
  "FROM tlkpSecType INNER JOIN tblSecurity ON tlkpSecType.lngSecTypeIdx = tblSecurity.lngSecTypeNum " & _
  "WHERE ([lngSecDealNum] = " & lngDeal & ") ORDER BY tblSecurity.lngSecDealNum, tblSecurity.lngSecTypeNum;"

  Set dbs = CurrentDb
  Set rst = dbs.OpenRecordset(strSQL, dbOpenDynaset)
  If rst.AbsolutePosition = -1 Then
    GoTo GetSecTypesStrExit
  End If
  rst.MoveLast
  rst.MoveFirst
  iTot = rst.RecordCount - 1
  
  Do While Not rst.EOF
    strType = rst(1)
    'Debug.Print "strType: " & strType
    If strType <> strPrevType Then 'process new entry
      If i = iTot Then
        str = str & strType
      Else
        str = str & strType & ", "
      End If
    End If
    strPrevType = strType
    
    rst.MoveNext
    i = i + 1
  Loop
  rst.Close
  dbs.Close
  
GetSecTypesStrExit:
  Set rst = Nothing
  Set dbs = Nothing
  GetSecTypesStr = str
  Exit Function
  
GetSecTypesStrErr:
  Debug.Print err.Number, err.Description
  Resume GetSecTypesStrExit
End Function
