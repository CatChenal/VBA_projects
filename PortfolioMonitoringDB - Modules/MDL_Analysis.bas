Attribute VB_Name = "MDL_Analysis"
Option Compare Database
Option Explicit
'================================================================================
'
' MDL_Analysis
'   Apr-6-04: Reviewed CompileAllDoc proc & sub-procs
'   Mar-15-04
'   Dec-19-03 10:00
'   Brought check on file existence early in the proc
'   Closed doc if not saved
'   Added ApplyTheme in Word to fix most final formatting problems
'
'================================================================================'
Public iPrevNodeIdx As Integer
Public Const cstFrontEndFolder = "H:\DB_FrontEnds\"
Public appWord As Word.Application

Const cstHTMLHeaderFile = "H:\Projects Active\Portfolio Monitoring\Missing Reviews-HEAD.txt"
Dim docMain As Word.Document
Dim objTree As TreeView
Dim intIndentLevel  As Integer
'
' QR=Quarterly Reviews. Path to store QRs:
Const cstQRFolderPath = "S:\PublicDB\Fundamentals\Mezzanine Quarterly Reviews\"
Const cstDefaultQRLocation = "S:\MezzInvestments\Closed\"
' PR=Published Reviews. Path to store compiled/published reviews:
Const cstPRFolderName = "AA Published Reviews" '& "\"
' To retrieve PR template: cstQRFolderPath & "\" & cstprFolderName & cstprTemplate
Const cstPRTemplate = "\Template\"
Const cstPRTemplateDoc = "Merged Reviews Template.dot"
'
Const cstQryCoNames = "qryCoNames"
Const cstFinPageNoFrc = "<<no forecast available>>"
Const cstFinPageNoData = "<<no data points for this period>>"
Const cstNoReview = " <Review not available> "
Const cstAdminGrp = "FundAdmin"
Const cstEndReview = " REVIEW END "
Const cstMainDocHdr = "MEZZANINE REVIEWS"
Const cstIndentSpaces = 2
Const cstLblCaption1 = "Add a completed review for this period:"
Const cstLblCaption2 = "Print above listed companies && their reviews"
'
Public Sub FillFileTree(intYear As Integer, iQ As Integer, _
                        Optional blnCreateMissingReviewsRpt As Boolean = False)
  Dim dbs As DAO.Database
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim frmReviews As Form
  Dim nodRoot As Node, nodCurrent As Node
  Dim fso As Scripting.FileSystemObject
  Dim fs As Scripting.TextStream
  Dim itmMail As Outlook.MailItem
  Dim strCo As String, strDocName As String, strDocPath As String, strPeriod As String
  Dim strRpt As String, strRecips As String, strAnalyst As String, strHeader As String
  Dim dte As Date
  Dim i As Integer
  Dim lngR As Long
  
  Const cstSubject = "Missing Quarterly Reviews"
  Const cstRptHdr = "<BODY><H2>Missing Quarterly Reviews</H2><BR><p>The following companies do not have a quarterly review ready for publication.  " & _
            "Please, use the button labeled <Add a completed review for this period> on the " & _
           "Analysis & Reviews page of the Portfolio Performance Monitor Database to add it " & _
            "to the completed reviews folder as soon as you have it ready. <br> Thank you. <br>" & _
            "</p> <p align=""center"">"

  Const cstTblHeader = "<TABLE> <TR> <TD class=""TDRowL"" > COMPANY </TD> " & _
                      "<TD class=""TDRowR"" > REVIEWER </TD> </TR><TR>"
  '------------------------------------------------------------------------------
  'On Error GoTo ProcErr
  Screen.MousePointer = 11 ''Busy (Hourglass)
  
  Set frmReviews = Forms(cstFRM_Main)!sfrmAny.Form
  If blnCreateMissingReviewsRpt And frmReviews!optAll Then
    MsgBox "The Missing Reviews Report is generated only when a quarterly period is selected.", _
            vbExclamation, "No Quarter Selected"
    GoTo ProcExit
  End If
  
  If Not blnCreateMissingReviewsRpt Then
    ' Refresh controls and tree according to action
    Call ShowPeriodCtls(frmReviews, Not frmReviews!optAll)
    Set objTree = frmReviews!ocxTree.Object
    objTree.Nodes.Clear
    ' Add top node (1). Here, it is just a title, not to be processed
    Set nodRoot = objTree.Nodes.Add(, , UCase(cstQRFolderPath), cstMainDocHdr)   '')
  Else
    strRpt = cstTblHeader 'prep email format
    ' Retrieve email html header:
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fs = fso.OpenTextFile(cstHTMLHeaderFile, ForReading)
    strHeader = fs.ReadAll
    fs.Close
    Set fso = Nothing: Set fs = Nothing
  End If
  
  ' Run qry for unrealized deals/cos in the selected quarter/year
  Set dbs = CurrentDb
  Set qdf = dbs.QueryDefs(cstQryCoNames)
  dte = GetGivenQtrDate(iQ, intYear, False) 'end dte of selected Q
  qdf.Parameters(0) = dte
  Set rst = qdf.OpenRecordset(dbOpenSnapshot)
  If rst.AbsolutePosition = -1 Then
    rst.Close
    qdf.Close
    GoTo ProcExit
  End If
  rst.MoveLast
  rst.MoveFirst

  If frmReviews!optAll Then
    strPeriod = "*.doc"
  Else
    strPeriod = "_" & intYear & "Q" & iQ
  End If

  With Application.FileSearch
    .FileType = 3 '= msoFileTypeWordDocuments
    For lngR = 1 To rst.RecordCount
      strCo = rst.Fields(0)  ' 0: txtName (=Co name); 1: txtUserName; 2: full name
      .LookIn = cstQRFolderPath & "\" & strCo & "\" 'Company folder for QReviews
      .FileName = strCo & strPeriod
  
      ' Set the current node to the current company
      If Not blnCreateMissingReviewsRpt Then
        ' Add the company folder node:
        Set nodCurrent = objTree.Nodes.Add(nodRoot, tvwChild, strCo, strCo)
      End If
      
      If .Execute() > 0 Then
        For i = 1 To .FoundFiles.Count
          strDocPath = .FoundFiles(i)
          strDocName = GetFileNameFromPath(strDocPath)
          If Not blnCreateMissingReviewsRpt Then
            'If tree node is that of a doc: node.key=path, node.text=short filename.
            objTree.Nodes.Add nodCurrent, tvwChild, strDocPath, strDocName
          End If
        Next i
      Else  ''.Execute() = 0
        If blnCreateMissingReviewsRpt Then

          strAnalyst = rst.Fields(2)
          If InStr(strRecips, strAnalyst) = 0 Then strRecips = strRecips & strAnalyst & ";"
          strRpt = strRpt & "<TD class=""TDRowL"">" & strCo & "</TD>" & _
                            "<TD class=""TDRowR"">" & strAnalyst & "</TD></TR>"
    
        End If
      End If
      rst.MoveNext
    Next lngR
    ' If email, close the html table w/proper tags:
    If blnCreateMissingReviewsRpt Then strRpt = strRpt & "</TABLE></p></BODY></HTML>"
  End With
  
  rst.Close
  qdf.Close
  Set nodCurrent = Nothing
  Set nodRoot = Nothing

  If blnCreateMissingReviewsRpt Then
    If Len(strRecips) = 0 Then
      MsgBox "There are no missing reviews for " & "Q" & iQ & " " & intYear, _
              vbInformation, "No Missing Reviews"
      GoTo ProcExit
    End If
    
    'Add descriptive header to msg:
    'strRpt = cstRptHdr & strRpt
    strRpt = strHeader & strRpt
    
    Call StartOutlook   ' : Start ol
    Set itmMail = ol.CreateItem(olMailItem)     ' : Create email
    With itmMail
      .To = strRecips
      .CC = cstAdminGrp
      .Subject = GetReviewTitle(frmReviews) & " - " & cstSubject
      'Debug.Print strRpt
      '.Body = strRpt
      '"H:\Projects Active\Portfolio MonitoringMissing Reviews-HEAD.txt" &
      .HTMLBody = strRpt
      .Display True
      .Recipients.ResolveAll
    End With
  Else
    Call ResetAnalysisPage
  End If
  
ProcExit:
  Set fso = Nothing: Set fs = Nothing
  Set rst = Nothing: Set qdf = Nothing: Set dbs = Nothing
  DoCmd.Hourglass False
  If blnCreateMissingReviewsRpt Then
    Set itmMail = Nothing
    Call EndOutlook
  End If
  Set nodRoot = Nothing
  Set nodCurrent = Nothing
  Set objTree = Nothing
  Set frmReviews = Nothing
  Exit Sub
  
ProcErr:
  If Err = 35602 Then 'key not unique
    Err.Clear
    Resume Next
  End If
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Proc: FillFileTree"
  Resume ProcExit
End Sub

Public Sub ShowPeriodCtls(frm As Form, blnShow As Boolean)
' Function called this way:  ShowPeriodCtls(frmReviews, Not frmReviews!optAll)
  On Error GoTo ProcErr
  DoCmd.Hourglass True
  With frm
    !lblShowPeriod.Visible = blnShow
    !lblAddReview.Visible = True
    If (-(Not blnShow) + 1) = 1 Then
      !lblAddReview.Caption = cstLblCaption1
    Else
     !lblAddReview.Caption = cstLblCaption2
    End If
    !cbxSelYear.Visible = blnShow
    !opgQ.Visible = blnShow
    !lblCompile.Visible = blnShow
    !lblMissingReviewRpt.Visible = blnShow
  End With
  DoCmd.Hourglass False
  Exit Sub
  
ProcErr:
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Proc: ShowPeriodCtls"
End Sub

Public Sub CompileAllDocs()
  Dim frmReviews As Form
  Dim strCo As String, strDocPath As String, strCompiledDoc As String
  Dim strTemplatePath As String, strPRPath As String, strDocTitle As String
  Dim f As Integer
  Dim blnSaveProp As Boolean, blnWordRunning As Boolean
  DoCmd.Hourglass True
  '------------------------------------------------------------------------------
  On Error GoTo ProcErr
  DoCmd.Hourglass True
  
  Set frmReviews = Forms(cstFRM_Main)!sfrmAny.Form
  strDocTitle = GetReviewTitle(frmReviews)
  If strLoggedUser = "" Then strLoggedUser = GetLoggedUser
  
  strPRPath = cstQRFolderPath & cstPRFolderName & "\"
  strCompiledDoc = strPRPath & strDocTitle & ".doc"
  If FileExists(strCompiledDoc) Then
    MsgBox strDocTitle & " already exists in the compiled & published reviews folder." & vbCrLf & _
          "Delete it or rename it before running the compilation." & vbCrLf & _
          "Published Reviews Folder = " & strPRPath, vbInformation, _
          "Compiled Document Already Exists"
    GoTo ProcExit
  End If
  
  Set appWord = AppOpen("OpusApp", "Word.Application", True, blnWordRunning)
  blnWordAlreadyRunning = blnWordAlreadyRunning Or blnWordRunning
  
  ' Open the MSWord & doc:
  strTemplatePath = strPRPath & cstPRTemplate & cstPRTemplateDoc
  Set docMain = appWord.Documents.Add(Template:=strTemplatePath)
  docMain.Activate
  docMain.ActiveWindow.ActivePane.View.ShowAll = False  'hide paragraph marks
  If docMain.Fields.Count > 0 Then
    For f = 1 To docMain.Fields.Count
      docMain.Fields.Item(f).Delete
    Next f
  End If
  docMain.Range(Start:=0, End:=0).Select
  appWord.Selection.ParagraphFormat.TabStops.Add Position:=InchesToPoints(7), _
                                                 Alignment:=wdAlignTabRight, Leader:=wdTabLeaderDots
  appWord.Selection.TypeParagraph
  appWord.Selection.MoveUp Unit:=wdLine, Count:=1

  Set objTree = frmReviews!ocxTree.Object
  ' Start with the first comp in the tree=node #2:
  Call InsertFile(objTree.Nodes(2))   ' recursively
  Call CreateTOC                      ' Table of contents
  Call InsertTitle(strDocTitle)       '
  Call ApplyTheme(docMain)            ' try clean up formating
     
  With docMain
    .BuiltinDocumentProperties(wdPropertySubject) = strDocTitle
    .BuiltinDocumentProperties(wdPropertyAuthor) = strLoggedUser
    blnSaveProp = .Application.Options.SavePropertiesPrompt
    .Application.Options.SavePropertiesPrompt = False
    .Application.Activate
    If MsgBox("Do you want to save the compiled document?", vbQuestion + vbYesNo, "Save?") = vbYes Then
      .Application.Options.SavePropertiesPrompt = blnSaveProp 'restore
      .SaveAs strCompiledDoc
    Else
      .Close SaveChanges:=acSaveNo
    End If
  End With

ProcExit:
  DoCmd.Hourglass False
  Set docMain = Nothing
  Set appWord = Nothing
  Set objTree = Nothing
  Set frmReviews = Nothing
  Exit Sub
  
ProcErr:
  If Err = 462 Then
    Err.Clear
    'Debug.Print "Err 462 when adding tab stop in CompileAllDocs"
    Resume Next
  Else
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Proc: CompileAllDocs"
    Resume ProcExit
  End If
End Sub

Sub InsertFile(objNode As Node)
  Dim strCo As String, strDisplay As String, strFile As String
  Dim blnFile As Boolean
  On Error GoTo ProcErr
  
  strFile = objNode.Key
  strDisplay = objNode.Text
  'if node is that of a company hdr then node.text = node.key
  blnFile = (strFile <> strDisplay) 'if true, then actual file name
  
  strCo = strDisplay
  If blnFile Then strCo = GetFileNameFromPath(strFile)
  If Not blnFile Then 'it's a header
    
    appWord.Selection.InsertBreak Type:=wdSectionBreakNextPage
    'Insert co name delimiter:
    appWord.Selection.TypeText Text:=UCase(strCo) & " REVIEW " & vbTab
    ' Insert TOC field:
    appWord.Selection.TypeParagraph
    appWord.Selection.Fields.Add Range:=appWord.Selection.Range, _
                                 Type:=wdFieldTOCEntry, _
                                 Text:=strCo, PreserveFormatting:=True
    If (objNode.Children = 0) Then
      appWord.Selection.TypeParagraph
      appWord.Selection.TypeText Text:=strCo & cstNoReview & vbTab
    Else
      InsertFile objNode.Child
    End If
    If TypeName(objNode.Next) <> "Nothing" Then InsertFile objNode.Next
    
  Else ' it's a file; insert
    appWord.Selection.TypeParagraph
    appWord.Selection.InsertFile FileName:=strFile, Range:="", ConfirmConversions:=False, _
                                 Link:=False, Attachment:=False
    appWord.Selection.TypeParagraph
    strCo = UCase(Left(strDisplay, InStrRev(strDisplay, "_") - 1))

    appWord.Selection.TypeText Text:=strCo & cstEndReview & vbTab
    appWord.Selection.TypeParagraph
  End If
  Exit Sub
  
ProcErr:
  DoCmd.Hourglass False
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Proc: InsertFile"
End Sub

Sub CreateTOC()
  On Error GoTo ProcErr ''Resume Next
  docMain.Range(Start:=0, End:=0).Select ' Go back to top of doc
  With appWord.Selection  ' Selection 'docMain ' Insert TOC:
    .Font.Size = 11
    .TypeText Text:="Table of Contents" & vbTab & "Page"
    .TypeParagraph
    .Font.Size = 8
  End With
  With appWord.ActiveDocument ' Insert TOC:
    .TablesOfContents.Add _
                      Range:=appWord.Selection.Range, _
                      RightAlignPageNumbers:=True, _
                      UseFields:=True, UseHeadingStyles:=False, _
                      IncludePageNumbers:=True, _
                      UseHyperlinks:=True
    .TablesOfContents(1).TabLeader = wdTabLeaderDots
    .TablesOfContents.Format = wdTOCFormal
  End With
  DoCmd.Hourglass False
  Exit Sub
  
ProcErr:
  DoCmd.Hourglass False
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Proc: CreateTOC"
End Sub

Sub InsertTitle(strTitle As String)
  On Error GoTo ProcErr
  
  docMain.Sections(1).Headers(wdHeaderFooterPrimary).Range.Select
  With appWord.Selection
    .MoveDown Unit:=wdLine, Count:=1
    .ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Font.Name = "Verdana"
    .Font.AllCaps = False
    .Font.Size = 11
    .TypeText Text:=strTitle & vbCrLf
  End With
  docMain.StoryRanges(wdMainTextStory).Select '(Start:=0, End:=0).Select
  docMain.Range(Start:=0, End:=0).Select
  Exit Sub
  
ProcErr:
  DoCmd.Hourglass False
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Proc: InsertTitle"
End Sub

Sub ApplyTheme(docActive As Word.Document)
  On Error GoTo ProcErr
 ' docActive.ApplyTheme Name:="Eclipse 000"
  With docActive.Application.ActiveWindow
    If .View.SplitSpecial = wdPaneNone Then
      .ActivePane.View.Type = wdPrintView
    Else
      .View.Type = wdPrintView
    End If
  End With
  With docActive
    .UpdateStylesOnOpen = True
    .AttachedTemplate = "S:\MezzInvestments\Deal Forms\Quarterly Review Template.dot"
  End With
  Exit Sub

ProcErr:
  DoCmd.Hourglass False
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Proc: CreateTOC"
End Sub

Public Function AddReviewToFolder()
' Output: 0=ok;  else: err.num
  Dim appMSWord As Word.Application
  Dim doc As Word.Document
  Dim frmReviews As Form
  Dim strCo As String, strPathFileToCopy As String, strTitle As String, strPeriod As String
  Dim strReviewNewName As String, strReviewFolder As String
  Dim intYear As Integer, intQ As Integer
  Dim lngResult As Long
  Dim blnPastReview As Boolean, blnWordRunning As Boolean
  Const cstExt = ".doc"
  DoCmd.Hourglass True
  On Error GoTo ProcErr
  
  Set frmReviews = Forms(cstFRM_Main)!sfrmAny.Form
  Set objTree = frmReviews!ocxTree.Object
  If TypeName(objTree.SelectedItem) = "Nothing" Then
    MsgBox "Highlight a Company in the list first.", vbInformation, "Select A Company"
    GoTo ProcExit
  End If
  
  ' Populate var
  strCo = frmReviews.Parent!cbxSelComp.Column(1)
  strReviewFolder = cstQRFolderPath & strCo & "\" '& cstqrFolderName & "\"
  If CreateReviewFolder(strReviewFolder) <> 0 Then GoTo ProcExit
  
  strPeriod = strCo & "_" & frmReviews!cbxSelYear & "Q" & frmReviews!opgQ & cstExt
  strReviewNewName = strReviewFolder & strPeriod
  
  If FileExists(strReviewNewName) Then
    ' Check if the review being added belongs to the current Q or beyond,
    ' then it's ok to overwrite if the file exists:
    blnPastReview = (frmReviews!cbxSelYear < Year(Date)) Or _
                    (frmReviews!opgQ < CInt(Format(Date, "q")))
    If blnPastReview Then
      MsgBox strPeriod & vbCrLf & "has already been archived and published. " & vbCrLf & _
            "You cannot overwrite any files in this folder." & vbCrLf & _
            "Please contact your database administrator for a workaround.", _
            vbExclamation, "No Past Reviews Overwriting Allowed"
      GoTo ProcExit
    End If
  End If
  
  strTitle = GetReviewTitle(frmReviews)
  strPathFileToCopy = ReturnOpenFileName(frmReviews, "*.doc" & Chr$(0) & "*.DOC", cstDefaultQRLocation, _
                                         "Locate " & strCo & "'s review for " & strTitle)
  If Len(strPathFileToCopy) = 0 Then
    MsgBox "You have not specified a file!" & vbCrLf & "Exiting...", vbExclamation, "No File Selected"
    GoTo ProcExit
  End If
 
  Set appWord = AppOpen("OpusApp", "Word.Application", True, blnWordRunning)
  blnWordAlreadyRunning = blnWordAlreadyRunning Or blnWordRunning

  Set doc = appMSWord.Documents.Open(FileName:=strPathFileToCopy, ReadOnly:=True, Visible:=True)
  With doc
    .Application.Options.SavePropertiesPrompt = True
    .BuiltinDocumentProperties(wdPropertyTitle) = strCo & " Quarterly Review"
    .BuiltinDocumentProperties(wdPropertySubject) = strTitle
    .SaveAs FileName:=strReviewNewName
    .Close SaveChanges:=wdPromptToSaveChanges
  End With
  appMSWord.WindowState = wdWindowStateMinimize
  If Not blnWordAlreadyRunning Then appMSWord.Quit
    
ProcExit:
  DoCmd.Hourglass False
  Set doc = Nothing
  Set appMSWord = Nothing
  Set objTree = Nothing
  Set frmReviews = Nothing
  AddReviewToFolder = lngResult
  Exit Function

ProcErr:
  lngResult = Err
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Proc: AddReviewToFolder"
  Resume ProcExit
End Function

Sub OpenPFMReview(strFileFullName As String, blnReadOnlyMode As Boolean)
  Dim strMsg As String
  Dim blnAppAlreadyRunning  As Boolean
  blnAppAlreadyRunning = False
  DoCmd.Hourglass True
  On Error GoTo OpenPFMReviewErr
  
  If blnReadOnlyMode Then
    strMsg = "You are about to open a Review in Read-Only mode because its creation " & _
             "quarter is no longer current." & vbCrLf & "Should you need to amend it, " & _
             "you will have to save this document under a new name."
    If MsgBox(strMsg, vbOKCancel + vbInformation, _
              "Open Past Quarter Review") = vbCancel Then GoTo OpenPFMReviewExit
  End If

  Set appWord = AppOpen("OpusApp", "Word.Application", True, blnAppAlreadyRunning)
  blnWordAlreadyRunning = blnWordAlreadyRunning Or blnAppAlreadyRunning
  
  With appWord
    .System.Cursor = wdCursorWait
    .Documents.Open FileName:=strFileFullName, ReadOnly:=blnReadOnlyMode
    .Activate
    .System.Cursor = wdCursorNormal
  End With
  
OpenPFMReviewExit:
  DoCmd.Hourglass False
  Set appWord = Nothing
  Exit Sub
  
OpenPFMReviewErr:
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "OpenPFMReview"
  Resume OpenPFMReviewExit
End Sub

Function GetReviewTitle(frm As Form) As String
  GetReviewTitle = Choose(frm!opgQ, "First", "Second", "Third", "Fourth") & " Quarter " & frm!cbxSelYear
End Function

Function CalcFinInfo(ByVal lngFrcID As Long, ByVal dteBeg As Date, ByVal dteEnd As Date)
  Dim dbDAO As DAO.Database
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim wbk As Excel.Workbook
  Dim wsh As Excel.Worksheet
  Dim varDebtFldsArray() As Variant
  Dim strMsg As String
  '
  Dim curNetEBITDA As Currency, curIntExp As Currency, curTotDebt As Currency
  Dim sglCoverage As Single, sglLeverage As Single
  Dim lngRows As Long, lng As Long
  Dim r As Integer, u As Integer, iMonth As Integer
  '
  Const cstDebtFieldsQry = "qryCurrentFrcDebtFields"
  Const cstEBITDAFieldsQry = "qryCurrentFrcEBITDAFields"
  Const cstIntExpFieldsQry = "qryCurrentFrcIntExpFields"
  Const cstFrcFreq = "qryCurrentFrcFrequency"
  Const cstBookName = cstFrontEndFolder & "FinInfo.xls"
  Const cstCov = "Coverage"
  Const cstLev = "Leverage"
  Const cstNetE = "Net EBITDA"
  Const cstIntX = "Int Exp"
  Const cstTotD = "Total Debt"
  Const cstPctDisp = "#.#x"
  Const cstListCol = 2
  Const cstAmtCol = 4
  '--------------------------------------------------------
  iMonth = 1:  r = 1 ' Initialize Row/Col numbers
  On Error GoTo CalcFinInfoErr
  DoCmd.Hourglass True
    
  Set dbDAO = CurrentDb
  '1. Get EBITDA Amts array
  Set qdf = dbDAO.QueryDefs(cstEBITDAFieldsQry)
  qdf.Parameters(0) = lngFrcID
  
  Set rst = qdf.OpenRecordset
  If rst.AbsolutePosition = -1 Then
    rst.Close
    qdf.Close
    strMsg = "No EBITDA Fields defined"
    GoTo CalcFinInfoExit
  End If
  Call GetFrcAmtsList(dbDAO, rst, lngFrcID, dteBeg, dteEnd, varDebtFldsArray, cstNetE)
  
  lngRows = 0
  lngRows = UBound(varDebtFldsArray, 1)
  curNetEBITDA = varDebtFldsArray(lngRows, 1) 'last row = total
  Erase varDebtFldsArray
  
  '2. Get IntExp Amts array
  Set qdf = dbDAO.QueryDefs(cstIntExpFieldsQry)
  qdf.Parameters(0) = lngFrcID
  
  Set rst = qdf.OpenRecordset
  If rst.AbsolutePosition = -1 Then
    rst.Close
    qdf.Close
    strMsg = "No Int. Exp Fields defined"
    GoTo CalcFinInfoExit
  End If
  Call GetFrcAmtsList(dbDAO, rst, lngFrcID, dteBeg, dteEnd, varDebtFldsArray, cstIntX)
  
  lngRows = 0
  lngRows = UBound(varDebtFldsArray, 1)
  curIntExp = varDebtFldsArray(lngRows, 1)  'last row = total
  Erase varDebtFldsArray
  
  '3. Get Tot Debt Amts array
  Set qdf = dbDAO.QueryDefs(cstDebtFieldsQry)
  qdf.Parameters(0) = lngFrcID
  
  Set rst = qdf.OpenRecordset
  If rst.AbsolutePosition = -1 Then
    rst.Close
    qdf.Close
    strMsg = "No Debt Fields defined"
    GoTo CalcFinInfoExit
  End If
  Call GetFrcAmtsList(dbDAO, rst, lngFrcID, dteBeg, dteEnd, varDebtFldsArray, cstTotD)
  
  lngRows = 0
  lngRows = UBound(varDebtFldsArray, 1)
  curTotDebt = varDebtFldsArray(lngRows, 1) 'last row = total
      
  '4. Get frc fequency (months)
  Set qdf = dbDAO.QueryDefs(cstFrcFreq)
  qdf.Parameters(0) = lngFrcID
  qdf.Parameters(1) = dteBeg

  Set rst = qdf.OpenRecordset
  If rst.AbsolutePosition = -1 Then
    rst.Close
    qdf.Close
    strMsg = "No period time frame (number of months) defined: defaulted to 1"
  Else
    rst.MoveFirst
    iMonth = rst(0)
    If iMonth = 0 Then
      strMsg = "Period time frame (number of months) found is 0: defaulted to 1"
      iMonth = 1
    End If
  End If
  rst.Close
  Set rst = Nothing
  qdf.Close
  Set qdf = Nothing
  
  ' Calculate ratios:
  If curIntExp <> 0 Then sglCoverage = curNetEBITDA / (curIntExp * iMonth)
  If curNetEBITDA <> 0 Then sglLeverage = curTotDebt / (curNetEBITDA * iMonth)

  ' Open xl book:
  Set wbk = GetPFMExcelBook(cstBookName, False) 'not shown;
  If wbk Is Nothing Then
    strMsg = "Error in GetPFMExcelBook: could not set object"
    GoTo CalcFinInfoExit
  End If
  wbk.Application.WindowState = xlMinimized
  Set wsh = wbk.Worksheets(1)
  With wsh.Range("Print_Area")
    .Cells.ClearContents
    With .Columns(cstAmtCol).Cells
      .Borders.LineStyle = xlLineStyleNone
      .Font.Bold = False
    End With
  
    ' Display result in stored spreadsheet (first col=margin column:blank)
    ' First 2 rows:
    .Cells(r, cstListCol) = UCase(varDebtFldsArray(0, 0))
    .Cells(r, cstAmtCol) = varDebtFldsArray(0, 1)
    r = r + 1
    .Cells(r, cstListCol) = UCase(varDebtFldsArray(1, 0))
    .Cells(r, cstAmtCol) = varDebtFldsArray(1, 1)
    r = r + 1
    .Cells(r, cstListCol) = "Amounts in $000's"
    r = r + 2
    
    'Cov & Lev:
    .Cells(r, cstListCol) = cstCov
    .Cells(r, cstAmtCol) = sglCoverage
    r = r + 1
    .Cells(r, cstListCol) = cstLev
    .Cells(r, cstAmtCol) = sglLeverage
    r = r + 2
    
    ' Others:
    .Cells(r, cstListCol) = cstNetE
    .Cells(r, cstAmtCol) = curNetEBITDA / iMonth
    r = r + 1
    .Cells(r, cstListCol) = cstIntX
    .Cells(r, cstAmtCol) = curIntExp / iMonth
    r = r + 1
    .Cells(r, cstListCol) = cstTotD
    .Cells(r, cstAmtCol) = curTotDebt / iMonth
    r = r + 2
    
    .Cells(r, cstListCol) = "DEBT FIELDS DETAILS:"
    r = r + 1
    
    For lng = 2 To lngRows
      .Cells(r, cstListCol) = varDebtFldsArray(lng, 0)
      .Cells(r, cstAmtCol) = varDebtFldsArray(lng, 1) / iMonth 'amounts
      If lng = lngRows Then u = r - 1
      r = r + 1
    Next lng
    .Cells(r - 1, cstAmtCol).Font.Bold = True ': bold total row
    
    With .Cells(u, cstAmtCol).Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .Weight = xlThin
    End With
  End With
  
  With Forms(cstFRM_Main)!sfrmAny.Form
    With !oleuOutput
      .Visible = True
      .Enabled = True
      .Locked = False
      .OLETypeAllowed = acOLEEither
      .Class = "Excel.Sheet.8"
      .SourceDoc = cstBookName
      .SourceItem = wsh.Range("Print_Area").Address(ReferenceStyle:=xlR1C1)
      .Action = acOLECreateLink
      .SizeMode = acOLESizeZoom
      .Locked = True
    End With
    !cbxSelPeriodTo.SetFocus
    !cbxSelPeriodTo.SelLength = 0
  End With
  wbk.Close SaveChanges:=True
  If Not blnExcelAlreadyRunning Then wbk.Application.Quit
  
CalcFinInfoExit:
  DoCmd.Hourglass False
  If Len(strMsg) > 0 Then MsgBox strMsg, vbExclamation, "CalcFinInfo"
  Set rst = Nothing
  Set qdf = Nothing
  Set wsh = Nothing
  Set wbk = Nothing
  Set dbDAO = Nothing
  Exit Function
  
CalcFinInfoErr:
  If Err = -2147417848 Then
    Err.Clear
    Resume Next
  Else
    MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, "CalcFinInfo"
    Resume CalcFinInfoExit
  End If
End Function

Function GetFrcAmtsList(dbDAO As DAO.Database, rstCoFields As DAO.Recordset, lngFRC As Long, _
                        dteFrom As Date, dteTo As Date, varOutputList As Variant, _
                        strAcctType As String) As Variant
'
  Dim rstCoAmts As DAO.Recordset
  Dim varTblFldNames() As Variant
  Dim strSQL As String, strSELECT As String, strFromWhere As String, strCrit As String, strLoopedSum As String
  Dim strTblFld As String, strFldShort As String, strPeriod As String
  Dim lngUBound As Long, lngFlds As Long, lngRecs As Long, f As Long
  Dim curTot As Currency, curAmt As Currency
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Const cstSELECT = "SELECT tblForecasts.txtForecastDesc AS Forecast, "
  Const cstFROMWHERE = " FROM tblForecasts INNER JOIN tblSeriesData ON tblForecasts.lngForecastID = " & _
                       "tblSeriesData.lngForecastID WHERE (((tblSeriesData.dtePeriodEndDate) "
  Const cstGROUPBY = " GROUP BY tblForecasts.txtForecastDesc, "
  Const cstAS = ") AS "
  Const cstRB = "[" 'right bracket
  Const cstLB = "]" 'left bracket
  Const cstCommaSum = ", Sum("
  ' Initialize ---------------------------------------------
  strPeriod = """" & dteFrom & " To " & dteTo & """"
  strSELECT = cstSELECT & strPeriod & " AS Period"
  strCrit = cstFROMWHERE & "Between #" & dteFrom & "# And #" & dteTo & "#) AND " & _
            "((tblSeriesData.lngForecastID)=" & lngFRC & ")) " & cstGROUPBY & strPeriod & " ;"
  strFromWhere = cstFROMWHERE & strCrit
  strLoopedSum = ", Sum("
  '----------------------------------------------------
                        
  With rstCoFields
  ' rstCoFields holds the company's fields: needed to recreate the query to list their amts.
    lngRecs = .RecordCount
    .MoveLast
    .MoveFirst
    'QRY FIELDS: 0:txtForecastDesc, 1:txtDispName, 2:txtFldTblName (, 3:txtAcctgCat: needed?)
    For f = 0 To lngRecs - 1
      strFldShort = cstRB & .Fields(1).Value & cstLB  'i.e.: AS [Term Loan A]
      strTblFld = cstRB & .Fields(2).Value & cstLB    'i.e.:[curTLA]
      strLoopedSum = strLoopedSum & strTblFld & cstAS & strFldShort
      If f <> lngRecs - 1 Then strLoopedSum = strLoopedSum & cstCommaSum
      .MoveNext
    Next f
  End With
  strSQL = strSELECT & strLoopedSum & strCrit

  ' Reset before reuse  -------------------------------
  f = 0: lngFlds = 0: strFldShort = ""
  '----------------------------------------------------
  
  Set rstCoAmts = dbDAO.OpenRecordset(strSQL)
  With rstCoAmts
    lngFlds = .Fields.Count
    'Size array and fill first column:
    ReDim varOutputList(lngFlds, 1)    ' Array of 2 cols and lngRecs+1 rows (extra row=total)-1
    For f = 0 To lngFlds - 1
      varOutputList(f, 0) = .Fields(f).Name
      If f > 1 Then
        curAmt = Nz(.Fields(f).Value, 0)
        varOutputList(f, 1) = curAmt
        curTot = curTot + curAmt
      Else
        varOutputList(f, 1) = .Fields(f).Value
      End If
    Next
  End With
  rstCoAmts.Close
  Set rstCoAmts = Nothing
  
  ' Populate last row of array:
  varOutputList(lngFlds, 0) = strAcctType
  varOutputList(lngFlds, 1) = curTot

  GetFrcAmtsList = varOutputList
  
End Function

Sub ResetAnalysisPage()
' Matches the company selected in the main form cbox to its entry in the tree list.
  Dim ocxTList As TreeView
  Dim nod As Node
  Dim str As String
  Dim blnReadOnly As Boolean
  Dim i As Integer
  On Error GoTo ProcErr
  
  With Forms(cstFRM_Main)
    ' Show "Read Only" lbl if file can no longer be edited (if opened in period < Q):
    blnReadOnly = (!dteQEndDate < GetGivenQtrDate(GetQtrFromDate(Date), Year(Date), False))
    !sfrmAny.Form!lblReadOnly.Visible = blnReadOnly

    !cbxSelComp.Requery
    If IsNull(!cbxSelComp.Column(1)) Then
      !cbxSelComp = !cbxSelComp.ItemData(0)
      !cbxSelForecast.Requery
    End If
    str = !cbxSelComp.Column(1)
    
    Set ocxTList = !sfrmAny.Form!ocxTree.Object
    On Error Resume Next
    Set nod = ocxTList.Nodes(str)
    If Err <> 0 Then
      If Err = 35601 Then 'not found
        Err.Clear
        ocxTList.SelectedItem = ocxTList.Nodes(1)
      Else
        GoTo ProcErr
      End If
    Else
      ocxTList.SelectedItem = ocxTList.Nodes(str)
    End If
  End With
  
  Call ResetFinancialPage
    
ProcExit:
  DoCmd.Hourglass False
  Set nod = Nothing
  Set ocxTList = Nothing
  Exit Sub
ProcErr:
  MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, "Proc: ResetAnalysisPage"
  Resume ProcExit
End Sub

Sub TreeClick()
' Synchronizes the main form selection box with the selected item in the tree
  Dim nod As Node
  Dim strCo As String
  Dim lngPos As Long
  
  'iPrevNodeIdx is set to 0 on form load
  With Forms(cstFRM_Main)!sfrmAny.Form
    If (!ocxTree.SelectedItem Is Nothing) Then Exit Sub
    If (!ocxTree.SelectedItem.Index = 1) Then Exit Sub
    If (!ocxTree.SelectedItem.Index = iPrevNodeIdx) Then Exit Sub
    
    Set nod = !ocxTree.SelectedItem
    iPrevNodeIdx = nod.Index
    strCo = nod.Text
    If (.Parent!cbxSelComp.Column(1) <> strCo) Then
    ' The selected item may not belong to the selected co in the main selection box
      lngPos = InStr(strCo, "_")
      If lngPos > 0 Then strCo = Left(strCo, lngPos - 1)
      .Parent!cbxSelComp.SetFocus
      .Parent!cbxSelComp.Text = strCo
      .Parent!cbxSelComp.Requery
      lngCurrentComp = .Parent!cbxSelComp.Column(0)
      .Parent!lblDefBud.Visible = False
      .Parent!cbxSelForecast.Requery
      !ocxTree.SetFocus
    End If
  End With
  Set nod = Nothing
End Sub

Sub ResetFinancialPage()
  Dim frmMain As Form, frm As Form
  Dim i As Integer
  Dim blnViewAll As Boolean
  On Error GoTo ProcErr
  
  Set frmMain = Forms(cstFRM_Main)
  Set frm = frmMain!sfrmAny.Form
  
  frmMain!lblDefBud.Visible = False
  frmMain!cbxSelForecast.Requery
  blnNoFrc = (frmMain!cbxSelForecast.ListCount = 0)
    
  blnViewAll = frm!optAll
  If frm!cbxSelPeriodFrom.Enabled Then frm!cbxSelPeriodFrom = Null
  If frm!cbxSelPeriodTo.Enabled Then frm!cbxSelPeriodTo = Null
  frm!cbxSelPeriodFrom.Enabled = Not blnNoFrc And Not blnViewAll
  frm!cbxSelPeriodTo.Enabled = Not blnNoFrc And Not blnViewAll
  frm!oleuOutput.Visible = False
  frm!lblTitlePeriod.Caption = ""
    
  If blnNoFrc Then
    frm!txtCurrentFrc = cstFinPageNoFrc
  Else
    ' Check if no frc selection, use first in list:
    frmMain!cbxSelForecast.SetFocus
    If (Len(frmMain!cbxSelForecast.Text) = 0) Then
      frmMain!cbxSelForecast = frmMain!cbxSelForecast.ItemData(0)
    End If
    frmMain!cbxSelForecast.SelLength = 0
    frmMain!lblDefBud.Visible = Nz(frmMain!cbxSelForecast.Column(3), 0)
    frm!txtCurrentFrc = frmMain!cbxSelForecast.Column(1)
    
    If Not blnViewAll Then  'specific period selected
      If frm!ocxTree.SelectedItem Is Nothing Then Exit Sub
      
      frm!cbxSelPeriodFrom = ""
      frm!cbxSelPeriodTo = ""
      frm!cbxSelPeriodFrom.Requery
      frm!cbxSelPeriodTo.Requery
       
      blnNoPoints = (frm!cbxSelPeriodFrom.ListCount = 0)
                
      frm!cbxSelPeriodFrom.Enabled = Not (blnNoFrc Or blnNoPoints)
      frm!cbxSelPeriodTo.Enabled = Not (blnNoFrc Or blnNoPoints)
      If blnNoPoints Then
        frm!lblTitlePeriod.Caption = cstFinPageNoData
      Else
        frm!cbxSelPeriodFrom = frm!cbxSelPeriodFrom.ItemData(0)
        i = frm!cbxSelPeriodTo.ListCount
        If i > 0 Then i = i - 1
        frm!cbxSelPeriodTo = frm!cbxSelPeriodTo.ItemData(i)
        
        frm!lblTitlePeriod.Caption = frm!cbxSelPeriodFrom & " To " & frm!cbxSelPeriodTo
      End If
    End If
  End If
  Set frm = Nothing
  Set frmMain = Nothing
    
  DoCmd.Hourglass False
  Exit Sub
  
ProcErr:
  Set frm = Nothing
  Set frmMain = Nothing
  DoCmd.Hourglass False
  MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, "Proc: ResetFinancialPage"
End Sub

Function CreateReviewFolder(strNewFolder As String)
  Dim fso As Scripting.FileSystemObject
  On Error Resume Next
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FolderExists(strNewFolder) Then
    fso.CreateFolder (strNewFolder)
    If Err <> 0 Then
      CreateReviewFolder = Err
      MsgBox "Error: " & Err & ", " & Err.Description, vbExclamation, "Proc: CreateReviewFolder"
      Resume Next
    End If
  End If
  Set fso = Nothing
End Function

Public Function FileExists(FileName As String) As Boolean   ' Return True if a file exists
  On Error GoTo FileExistsErr
  ' get the attributes and ensure that it isn't a directory
  If (GetAttr(FileName) And vbDirectory) = 0 Then FileExists = True
  Exit Function
FileExistsErr:
  FileExists = False
End Function

'====== UTIL PROC ================================================
Function GetDocList(strList As String)
  Dim frmReviews As Form
  Set frmReviews = Forms(cstFRM_Main)!sfrmAny.Form 'Analysis page has to be loaded!
  Set objTree = frmReviews!ocxTree.Object
  intIndentLevel = 1
  Call ParseTree(objTree.Nodes(1), strList)
  Set objTree = Nothing
End Function

'====== UTIL PROC ================================================
Sub RunGetDocList()
  Dim strOut As String
  Call GetDocList(strOut)
  Debug.Print strOut
End Sub

'====== UTIL PROC ================================================
Sub ParseTree(objNode As Node, strOutput As String)
  ' Print the node that was passed in and account for the node's level
  strOutput = strOutput & Space(intIndentLevel * cstIndentSpaces) & objNode.Text & vbCrLf
  
  ' Check to see if the current node has children
  If objNode.Children > 0 Then
    ' Increase the indent if children exist
    intIndentLevel = intIndentLevel + 1
    ' Pass the first child node to the print routine
    Call ParseTree(objNode.Child, strOutput)
    strOutput = strOutput & vbCrLf
  End If
  ' Set the next node to print
   Set objNode = objNode.Next
   ' As long as we have not reached the last node in a branch, continue to call the print routine
   If TypeName(objNode) <> "Nothing" Then
     Call ParseTree(objNode, strOutput)
   Else
     ' If the last node of a branch was reached, decrease the indentation counter
     intIndentLevel = intIndentLevel - 1
   End If
End Sub
