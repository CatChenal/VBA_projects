Attribute VB_Name = "Util-Tbl-Qry"
Option Compare Database
Option Explicit
'
'*****************************************************************
' DF Reports-Util-Tbl-Qry 10-01-03 16:50
'******************************************************************
'

Function ListQtrQueries(lngYr As Long, iQ As Integer)
  Dim qdf As DAO.QueryDef
  Dim strYrFind As String, strQFind As String, strPartialQryName As String
  Dim i As Integer

  i = 0
  Debug.Print "#    Query name:"
  strYrFind = "qry" & lngYr   'qry2001
  For Each qdf In CurrentDb.QueryDefs
    If Left(qdf.Name, 7) = strYrFind Then
      strPartialQryName = Mid$(qdf.Name, 8)
      strQFind = "Q" & iQ
      If InStr(strPartialQryName, strQFind) > 0 Then
        i = i + 1
        Debug.Print Format(i, "00") & ":  " & qdf.Name
      End If
    End If
  Next qdf
  qdf.Close
  Set qdf = Nothing

End Function

Function ListYrQueries(lngYr As Long)
  Dim qdf As DAO.QueryDef
  Dim strQry As String, strQryFile As String
  Dim i As Integer, iFileNum As Integer
  Dim varPos As Variant
  Const cstQryFile = "H:\Projects Active\Deal Flow Database (DFD)\DFD Development copies\"
  
  On Error GoTo ListYrQueriesErr
  iFileNum = FreeFile
  strQryFile = cstQryFile & "YrQryList" & lngYr & ".txt"

  i = 0
  Open strQryFile For Output As #iFileNum
  For Each qdf In CurrentDb.QueryDefs
    strQry = qdf.Name
    If Left(strQry, 7) = CStr("qry" & lngYr) Then
      i = i + 1
      If (InStr(8, strQry, "Q1") = 0) And (InStr(8, strQry, "Q2") = 0) And _
         (InStr(8, strQry, "Q3") = 0) And (InStr(8, strQry, "Q4") = 0) Then
        'Debug.Print strQry
        Write #iFileNum, i & ": " & strQry & vbCrLf & qdf.SQL & vbCrLf
      End If
    End If
  Next qdf
  qdf.Close
  Close #iFileNum
  Beep
  
ListYrQueriesExit:
  Set qdf = Nothing
  Exit Function
    
ListYrQueriesErr:
  MsgBox err.Number
  Resume ListYrQueriesExit
End Function

Function ListQueriesSQL(lngYr As Long, iQtr As Integer)
  Dim qdf As DAO.QueryDef
  Dim strQry As String, strTextBody As String, strQryFile As String
  Dim strCompare As String, strQ As String
  
  Dim i As Integer, iFileNum As Integer
  Dim varPos As Variant
  Const cstQryFile = "H:\Projects Active\Deal Flow Database (DFD)\DFD Development copies\"
  strTextBody = ""
  
  On Error GoTo ListQueriesSQLErr
  
  If MsgBox("List queries SQL?", vbYesNo) = vbYes Then
    iFileNum = FreeFile

    If iQtr > 0 Then
      strQryFile = cstQryFile & "QryList" & lngYr & "-Q" & iQtr & ".txt"
    Else
      strQryFile = cstQryFile & "YrQryList" & lngYr & ".txt"
    End If
    
    'Debug.Print strQryFile & vbCrLf & "Qry count: " & dbs.QueryDefs.Count
    i = 0
    Open strQryFile For Output As #iFileNum
    For Each qdf In CurrentDb.QueryDefs
      strQry = qdf.Name
      strCompare = CStr("qry" & lngYr)
      strTextBody = ""
      
      If Left(strQry, 7) = strCompare Then
        i = i + 1
        strQ = "Q" & iQtr
        If InStr(8, strQry, strQ) = 0 Then
          Debug.Print strQry
          strTextBody = i & ": Qry " & strQry & ": " & vbCrLf & qdf.SQL & vbCrLf
          Write #iFileNum, strTextBody
        End If
      End If
    Next qdf
    qdf.Close
    Close #iFileNum
    
  End If
  
ListQueriesSQLExit:
  Set qdf = Nothing
  Exit Function
    
ListQueriesSQLErr:
  MsgBox err.Number
  Resume ListQueriesSQLExit
End Function

Function ListQueriesAsFunctions(lngYr As Long, iQtr As Integer)
  Dim qdf As DAO.QueryDef
  Dim strQry As String
  Dim i As Integer
  
  If MsgBox("List queries as functions?", vbYesNo) = vbYes Then
    For Each qdf In CurrentDb.QueryDefs
      strQry = qdf.Name
      If Left(strQry, 7) = "qry" & lngYr Then 'function SQLxx(lngYr as long, iQtr as integer) as string '
        i = i + 1
        If i > 59 Then Exit Function
        If iQtr > 0 Then
          If Not IsNull(InStr(9, strQry, "Q" & iQtr)) Then
            Debug.Print "function SQL_" & Mid$(strQry, 8) & _
                        "(lngYr as long, iQtr as integer) as string ' " & i & vbCrLf & _
                        "end function" & vbCrLf
          End If
        ElseIf iQtr = 0 Then 'ignore q number
            Debug.Print "function SQL_" & Mid$(strQry, 8) & "(lngYr as long, iQtr as integer) as string ' " & i & vbCrLf & _
                        "end function" & vbCrLf
        
        End If
      End If
    Next qdf
    qdf.Close
  End If
  
  If MsgBox("List functions name?", vbYesNo) = vbYes Then
    ' List functions name:
    For Each qdf In CurrentDb.QueryDefs
      strQry = qdf.Name
      If Left(strQry, 7) = "qry" & lngYr Then  'function SQLxx(lngYr as long, iQtr as integer) as string '
        i = i + 1
        Debug.Print "SQL_" & Mid$(strQry, 8)
      End If
    Next qdf
    qdf.Close
  End If
  Set qdf = Nothing
  
End Function

Public Function ListQryParams(strQdfinCurrentDB As String)
  Dim qdf As DAO.QueryDef
  Dim param As DAO.Parameter
  Dim fld As DAO.Field
  Dim strHdr As String ', strType As String
  Dim p As Integer, i As Integer
   i = 0: p = 0
     
   On Error Resume Next
   Set qdf = CurrentDb.QueryDefs(strQdfinCurrentDB)
   
   If err.Number <> 0 Then
      If err.Number = 3265 Then
         MsgBox "This query: " & strQdfinCurrentDB & " does not exists or its " & _
                                         "name is mispelled"
         Set qdf = Nothing
         Exit Function
      Else
         MsgBox "Error: " & err.Description & " - " & err.Number
         Set qdf = Nothing
         Exit Function
      End If
   Else
      p = qdf.Parameters.Count
      If p = 0 Then
         Debug.Print " This query does not use parameters, or they are not accessible."
      Else
         strHdr = vbCr & "The query '" & strQdfinCurrentDB & "' has " & p & " parameter(s):"
         Debug.Print strHdr; Tab; Tab; "(" & Format(Date, "Medium Date") & ")" & vbCr
         For i = 0 To p - 1
            Set param = qdf.Parameters(i)
            Debug.Print Tab; Format(i, "00") & " :  " & param.Name
        Next i
      End If
   End If

  qdf.Close
  Set param = Nothing
  Set qdf = Nothing

End Function

Public Function ListQryFields(inQdf As String)
  Dim qdf As DAO.QueryDef
  Dim fld As DAO.Field
  Dim strHdr As String, strType As String
  Dim i As Integer
  Dim l As Long
  i = 0: l = 0
  strHdr = "": strType = ""
  On Error Resume Next
   
   Set qdf = CurrentDb.QueryDefs(inQdf)
   If err.Number <> 0 Then
      If err.Number = 3265 Then
         MsgBox "This query: " & inQdf & " does not exists or its " & _
                                         "name is mispelled"
         Set qdf = Nothing
         Exit Function
      Else
         MsgBox "Error: " & err.Description & " - " & err.Number
         Set qdf = Nothing
         Exit Function
      End If
   Else
      strHdr = vbCrLf & "Fields in the query: " & qdf.Name & vbCrLf & _
               "In the database: " & CurrentDb.Name
      Debug.Print strHdr; Tab; Date & vbCrLf
      
      For Each fld In qdf.Fields
         strType = ""
         i = i + 1
         Select Case fld.Type
            Case 3
               strType = "dbInteger"
            Case 4
               strType = "dbLong" ' or dbAutoNumber"
            Case 5
               strType = "dbCurrency"
            Case 7
               strType = "dbDouble"
            Case 8
               strType = "dbDate"
            Case 10
               strType = "dbText"
            Case 12
               strType = "dbMemo"
            Case 15
               strType = "dbGUID"
            Case Else
               strType = "undetermined: " & fld.Type
         End Select
         l = 20 - Len(fld.Name)
         Debug.Print Format(i, "00") & " :  " & fld.Name; Spc(l); Tab; strType
      Next fld
   End If
   qdf.Close
   Set fld = Nothing
   Set qdf = Nothing
   
End Function

Public Function ListTBLFields(inTdf As String)
  Dim tdf As DAO.TableDef
  Dim fld As DAO.Field
  Dim i As Integer
  Dim l As Long, m As Long
  Dim varDesc As Variant
  Dim strFields As String, strHdr As String, strType As String, strDesc As String
  Dim strDefVal As String
   
  i = 0: strFields = ""
  On Error Resume Next
  
  Set tdf = CurrentDb.TableDefs(inTdf)
  If err.Number <> 0 Then
     If err.Number = 3265 Then
        MsgBox "This table: " & inTdf & " does not exists or its " & _
                                        "name is mispelled"
        Set tdf = Nothing
        Exit Function
     Else
        MsgBox "Error: " & err.Description & " - " & err.Number
        Set tdf = Nothing
        Exit Function
     End If
  Else
     If tdf.Connect <> "" Then
       strHdr = vbCrLf & "Fields in the table named: '" & tdf.Name & _
                "'" & vbCrLf & _
                "Belonging to: '" & CurrentDb.Name & "'" & vbCrLf & _
                "Linked to: '" & tdf.Connect & "'."
     Else
       strHdr = vbCrLf & "Fields in the table named: '" & tdf.Name & _
                "'" & vbCrLf & "Belonging to: '" & CurrentDb.Name & "'"
     End If
     Debug.Print strHdr; Tab; Tab; "List date: " & Now & vbCrLf
     
     For Each fld In tdf.Fields
        varDesc = Null
        strDesc = ""
        strType = ""
        i = i + 1
        Select Case fld.Type
           Case 15
              strType = "dbGUID"
           Case 1
             strType = "dbBoolean"
           Case 3
              strType = "dbInteger"
           Case 4
              strType = "dbLong"
           Case 5
              strType = "dbCurrency"
           Case 6
             strType = "dbSingle"
           Case 7
              strType = "dbDouble"
           Case 8
              strType = "dbDate"
           Case 10
              strType = "dbText"
           Case 12
              strType = "dbMemo"
           Case Else
              strType = "undetermined: " & fld.Type
        End Select
        varDesc = fld.Properties("Description")
        strDesc = Nz(varDesc, "(no desc.)")
        strDefVal = IIf(IsNull(fld.DefaultValue) Or fld.DefaultValue = "", "", ": " & fld.DefaultValue)
        l = 20 - Len(fld.Name)
        m = 24 - Len(strType)
        Debug.Print Format(i, "00") & " : " & fld.Name; Tab; Tab; strType; Tab; strDesc; Space(m); Tab; strDefVal
     Next fld
  End If
  Set fld = Nothing
  Set tdf = Nothing
   
End Function

Sub CountFields(strTBL As String)
  Dim tdf As DAO.TableDef
  Dim fld As DAO.Field

  Set tdf = CurrentDb.TableDefs(strTBL)
  Debug.Print tdf.Fields.Count
  For Each fld In tdf.Fields
     Debug.Print fld.Name
  Next fld
  Set fld = Nothing
  Set tdf = Nothing
End Sub

Public Function ListCurrentDBTables()
  Dim dbs As DAO.Database
  Dim tbl As DAO.TableDef
  Dim strHdr As String, strDesc As String, strConnect As String
  Dim strDrive As String, strDB As String, strLocalDrive As String
  Dim i As Integer, j As Integer
  Dim l As Long, m As Long, pos As Long

  i = 0
  Set dbs = CurrentDb
  strLocalDrive = Left(dbs.Name, 1)
  strHdr = "Table(s) in " & dbs.Name & " as of " & Format(Date, "Medium Date") & vbCr
  Debug.Print strHdr
  Debug.Print "#    Table name:"; Spc(10); "Description:"; Spc(27); "Linked/Local: Drive"
  Debug.Print String(80, "-")
  
  For Each tbl In dbs.TableDefs
     strConnect = ""
     strDesc = ""
     If Left(tbl.Name, 4) <> "MSys" Then
        i = i + 1
        For j = 0 To tbl.Properties.Count - 1
           If tbl.Properties(j).Name = "Description" Then
              strDesc = tbl.Properties(j).Value
           End If
        Next j
        strConnect = tbl.Connect
        If tbl.Connect <> "" Then
           pos = InStr(1, strConnect, ";")
           strDB = Mid$(strConnect, pos + 10)
           strDrive = Left(strDB, 1)
           strConnect = "Linked"
        Else
           strConnect = "LOCAL"
           strDrive = strLocalDrive
        End If
        l = 20 - Len(tbl.Name)
        m = 40 - Len(strDesc)
        Debug.Print Format(i, "00") & ":  " & tbl.Name; Spc(l); strDesc; Spc(m); strConnect & ": " & strDrive
     End If
  Next tbl
  Set tbl = Nothing
  dbs.Close
  Set dbs = Nothing
End Function

Public Function ListCurrentDBQueries()
  Dim dbs As DAO.Database
  Dim qdf As DAO.QueryDef
  Dim qdfProp As DAO.Property
  Dim strHdr As String, strDesc As String
  Dim i As Integer, j As Integer
  Dim l As Long, m As Long

  i = 0
  Set dbs = CurrentDb
  strHdr = "Table(s) in " & dbs.Name & " as of " & Format(Date, "Medium Date") & vbCr
  Debug.Print strHdr
  Debug.Print "#    Query name:"; Spc(10); "Description:"
  Debug.Print String(80, "-")
   
  For Each qdf In dbs.QueryDefs
     strDesc = ""
     If Left(qdf.Name, 1) <> "Z" Then
     ' Queries starting with Z are under contruction or for Admin use.
        i = i + 1
        For j = 0 To qdf.Properties.Count - 1
           If qdf.Properties(j).Name = "Description" Then
              strDesc = qdf.Properties(j).Value
           End If
        Next j
        l = 20 - Len(qdf.Name)
        m = 40 - Len(strDesc)
        Debug.Print Format(i, "00") & ":  " & qdf.Name; Spc(l); strDesc
     End If
  Next qdf
  Set qdf = Nothing
  dbs.Close
  Set dbs = Nothing
End Function

Public Function RefreshTblLinks()
'--------------------------------------------------------------------
' SPECS: If the connect string of a linked table contains a
'        drive letter that is different form that of the front end,
'        all the linked table are refreshed to the FE location.
'
'--------------------------------------------------------------------
  Dim tdf As DAO.TableDef
  Dim strFrontDrive As String, strBackDrive As String, strConnect As String, strDB As String
  Dim i As Integer
  Dim pos As Long
  Const cstSharedBackEndConnect = ";DATABASE=S:\Databases\DealLog\DFDB2000_BE.mdb"
  Const cstDevelBackEndConnect = ";DATABASE=H:\Projects\Deal Flow Database (DFD)\DFD Development copies\DFDB2000_BE.mdb"
  i = 0
  On Error GoTo RefreshTblLinksErr
  DoCmd.Hourglass True
  
  strFrontDrive = Left(CurrentDb.Name, 1)
  
  For Each tdf In CurrentDb.TableDefs
    strConnect = ""
    strConnect = tdf.Connect
    'Sample connect string:
    ';DATABASE=H:\Projects\Deal Flow Database (DFD)\DFD Development copies\DFDB2000_BE.mdb
    If strConnect <> "" Then  'table is linked
      'find if backend db on same drive (current setup 9/25)
      pos = InStr(1, strConnect, ";")
      strDB = Mid$(strConnect, pos + 10)  'Len("DATABASE=")=9
      strBackDrive = Left(strDB, 1)
      If strFrontDrive <> strBackDrive Then
        i = i + 1
        Select Case strFrontDrive
          Case "H"
            tdf.Connect = cstDevelBackEndConnect
          Case "S"
            tdf.Connect = cstSharedBackEndConnect
        End Select
        tdf.RefreshLink
      End If
    End If
  Next tdf
  Set tdf = Nothing
  If i = 0 Then Debug.Print "No table reconnection needed"
  DoCmd.Hourglass False
  
RefreshTblLinksExit:
  Set tdf = Nothing
  Exit Function
  
RefreshTblLinksErr:
  DoCmd.Hourglass False
  MsgBox "Error: " & err.Number & " : " & err.Description, vbExclamation, "RefreshTblLinks"
  Resume RefreshTblLinksExit
End Function

Sub AdHocProcForQueries()
  Dim qdf As QueryDef
  Dim strQry As String, strOut As String
  Dim Q As Integer
  Dim lngYr As Long
  Dim blnDelAll As Boolean, blnLog As Boolean
  Const cstLogPath = "H:\Projects Active\Deal Flow Database (DFD)\AdHocProc-Output.txt"

  strOut = "": lngYr = 0: Q = 0: blnDelAll = False: blnLog = False
  
  lngYr = CLng(InputBox(vbCrLf & "Enter the Year:", "AdHocProcForQueries"))
  If lngYr < 1998 Then Exit Sub
  
  If MsgBox("List queries?", vbYesNo) = vbYes Then
    If MsgBox("Log list?", vbYesNo) = vbYes Then blnLog = True
    If MsgBox("Delete all?", vbYesNo) = vbYes Then blnDelAll = True
    
    For Each qdf In Application.CurrentDb.QueryDefs
      strQry = qdf.Name
      If Left(strQry, 7) = "qry" & lngYr Then 'function SQLxx(lngYr as long, iQtr as integer) as string '
        Q = Q + 1
        'strOut = strOut & "Function SQL_" & Mid$(strQry, 4) & "(lngYr as long, iQtr as integer) as string ' " & Q & vbCrLf & _
                vbTab & "Dim strSQL as string" & vbCrLf & "strSQL = " & qdf.SQL & vbCrLf & _
                "End function" & vbCrLf
        'strOut = strOut & "' SQL_" & Mid$(strQry, 4) & vbCrLf
        If blnLog Then strOut = strOut & strQry & vbCrLf
        If blnDelAll Then Application.CurrentDb.QueryDefs.Delete strQry
        
      End If
    Next qdf
    If blnLog Then Call SaveToFile(cstLogPath, strOut, True)
  End If
  Set qdf = Nothing
  
  MsgBox "AdHocProcForQueries Over"
End Sub

Function SaveToFile(FileName$, Body$, blnOverWriteExisting As Boolean)
  Dim fso As Object
  Dim fsFile As Object
  
  On Error GoTo SaveToFileErr
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set fsFile = fso.CreateTextFile(FileName$, blnOverWriteExisting)
  fsFile.WriteLine (Body$)
  fsFile.Close
  
SaveToFileExit:
  Set fsFile = Nothing
  Set fso = Nothing
  Exit Function
  
SaveToFileErr:
  MsgBox err.Number
  Resume SaveToFileExit
End Function
