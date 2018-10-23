Attribute VB_Name = "InsertMDBQuery"
Option Explicit
' Last update: 4/5/2004 1:48:22 PM
'
Sub InsertAccessQry()
  Dim varConnect As Variant
  Dim strSQL As String, strQryName As String, strRange As String, strFullQryName As String
  Dim strDestinationCell As String
  Dim rngData As Range
  Const cstDefDir = "S:\Databases\DealLog\"
  Const cstSharedDB = "DF Reports.mdb"
  '===================================================================================
  On Error GoTo InsertAccessQryErr
  
  strQryName = InputBox("Query to link to:", "Name of the Access qry to link to", "")
  If Len(strQryName) = 0 Then
    MsgBox "No entry: exiting function", , "Action Canceled"
    Exit Sub
  End If
  'Debug.Print "strQryName: " & strQryName
  
  If Selection Is Nothing Then
    strDestinationCell = InputBox("Enter the destination range:", "Output cell or range", "")
    If Len(strDestinationCell) = 0 Then
      MsgBox "No entry: exiting function", , "Action Canceled"
      Exit Sub
    End If
    Range(strDestinationCell).Select
  End If

  ' Clear selection data (keep headers)
  strRange = Selection.Cells(2, 1).Address & ":" & _
       Selection.Cells(Selection.Rows.Count, Selection.Columns.Count).Address
  Range(strRange).ClearContents
  ' Use first cell of range for destination
  strDestinationCell = Selection.Cells(1, 1).Address
 
  varConnect = Array( _
                  Array("ODBC;DSN=MS Access Database;" & _
                    "DBQ=" & cstDefDir & cstSharedDB & ";" & _
                    "DefaultDir=" & cstDefDir & ";" & _
                    "DriverId=25;FIL=MS Access;MaxBuf"), _
                  Array("ferSize=2048;PageTimeout=5;"))
 
  strFullQryName = "[" & cstDefDir & cstSharedDB & "].[" & strQryName & "]"
 
  If strQryName = "qryClosedDealsSummary" Then 'pick fields to return
    strSQL = "SELECT [Yr Closed], [Q Closed], Issuer, Source, [Deal Type], " & _
              "Coverage, Leverage, Securities " & Chr(13) & "" & Chr(10) & _
              "FROM " & strFullQryName
  Else
    strSQL = "SELECT * " & Chr(13) & "" & Chr(10) & "FROM " & strFullQryName
  End If
  'Debug.Print "strSQL: " & strSQL
  
  With ActiveSheet.QueryTables.Add(Connection:=varConnect, Destination:=Range(strDestinationCell))
      .CommandText = Array(strSQL)
      .CommandType = xlCmdSql
      .Name = strQryName
      .FieldNames = True
      .RowNumbers = False
      .FillAdjacentFormulas = False
      .PreserveFormatting = True
      .RefreshOnFileOpen = False
      .BackgroundQuery = True
      If strQryName = "qryClosedDealData" Then
          'Or strQryName = "qryClosedDealsSummary" Then
        .RefreshStyle = xlInsertEntireRows
      Else
        .RefreshStyle = xlOverwriteCells
      End If
      .SavePassword = True
      .SaveData = True
      .AdjustColumnWidth = True
      .RefreshPeriod = 0
      .PreserveColumnInfo = True
      .Refresh
  End With

InsertAccessQryExit:
  Exit Sub
  
InsertAccessQryErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, "InsertAccessQry"
  Resume InsertAccessQryExit
End Sub

