VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
' DBFrontEnd
' CLSMDL frmDeal Sep-05-03 14:40
'
'================================================================================
Const cstThisForm = cstDealForm
'-----------------------------------------------------------------------------
  
Private Sub cbxDealAnalyst_BeforeUpdate(Cancel As Integer)
  Cancel = ChangeOfAnalyst(Screen.ActiveControl)
End Sub

Private Sub cbxDealAnalyst2_BeforeUpdate(Cancel As Integer)
  Cancel = ChangeOfAnalyst(Screen.ActiveControl)
End Sub

Private Sub cbxDealAnalystMarket_BeforeUpdate(Cancel As Integer)
  Cancel = ChangeOfAnalyst(Screen.ActiveControl)
End Sub

Private Sub cbxDealAnalyst2_AfterUpdate()
  Me!lngDealSecondAnalNum = Me!cbxDealAnalyst2.Column(0)
End Sub

Private Sub cbxDealAnalyst_AfterUpdate()
  With Me
    If Not !cbxDealAnalyst.Column(2) Then Exit Sub
 
    !lngDealAnalNum = !cbxDealAnalyst.Column(0) 'set analyst num to column with num
    If IsNull(!cbxDealAnalyst2) Then
      !cbxDealAnalyst2 = !cbxDealAnalyst
      Call cbxDealAnalyst2_AfterUpdate
    End If
    If IsNull(!cbxDealAnalystMarket) Then
      !cbxDealAnalystMarket = !cbxDealAnalyst
      Call cbxDealAnalystMarket_AfterUpdate
    End If
  End With
End Sub

Private Sub cbxDealAnalystMarket_AfterUpdate()
  Me!lngMarketingAnalNum = Me!cbxDealAnalystMarket.Column(0)
End Sub

Private Sub cbxDealDispo_AfterUpdate()
  With Me
    If !dteDealDateIn = Date Then !dteDealDateDisp = Date
    !cbxSelRole = 4
    !cbxSelRole.SetFocus
  End With
End Sub

Private Sub cbxDealIssuer_AfterUpdate()
   Call DealEntitiesAfterUpd(Me!cbxDealIssuer)
End Sub

Private Sub cbxDealSourceName_AfterUpdate()
  Call DealEntitiesAfterUpd(Me!cbxDealSourceName)
End Sub

Private Sub cbxEqSponsor_AfterUpdate()
  Call DealEntitiesAfterUpd(Me!cbxEqSponsor)
End Sub

Private Sub cbxDealIssuer_BeforeUpdate(Cancel As Integer)
  Set cbxCallingBox = Me!cbxDealIssuer
End Sub

Private Sub cbxDealSourceName_BeforeUpdate(Cancel As Integer)
  Set cbxCallingBox = Me!cbxDealSourceName
End Sub

Private Sub cbxEqSponsor_BeforeUpdate(Cancel As Integer)
  Set cbxCallingBox = Me!cbxEqSponsor
End Sub

Private Sub cbxDealSourceName_NotInList(NewData As String, Response As Integer)
  Response = DealEntityNotInList(Me!cbxDealSourceName, NewData)
End Sub

Private Sub cbxDealIssuer_NotInList(NewData As String, Response As Integer)
  Response = DealEntityNotInList(Me!cbxDealIssuer, NewData)
End Sub

Private Sub cbxEqSponsor_NotInList(NewData As String, Response As Integer)
  Response = DealEntityNotInList(Me!cbxEqSponsor, NewData)
End Sub

Private Sub cbxDealStage_AfterUpdate()
  If Me!dteDealDateIn < Date Then Exit Sub
  Me!dteDealStageDate = Date
End Sub

Private Sub cmbPrintCurrent_Click()
  DoCmd.OpenReport "RPT_Current Deal Data", acPreview
End Sub

Private Sub dteDealDateDisp_DblClick(Cancel As Integer)
  Me!dteDealDateDisp = Date
End Sub

Private Sub dteDealDateIn_DblClick(Cancel As Integer)
  Me!dteDealDateIn = Date
End Sub

Private Sub dteDealDateIn_AfterUpdate()
  With Me
    !dteDealDateDisp = !dteDealDateIn
    !dteDealStageDate = !dteDealDateIn
  End With
End Sub

Private Sub dteDealStageDate_DblClick(Cancel As Integer)
  Me!dteDealStageDate = Date
End Sub

Private Sub Form_Activate()
  DoCmd.Maximize
  Call ResetToolbars
  Call ApplyFormFilter(cstThisForm)
    
  If TypeName(cbxCallingBox) <> "Nothing" Then
    cbxCallingBox.SetFocus
    cbxCallingBox.Requery
    Set cbxCallingBox = Nothing
  End If
  If blnRequeryCbx Then
    With Me
      !cbxDealSourceName.Requery
      !cbxEqSponsor.Requery
      !cbxDealIssuer.Requery
    End With
  End If
End Sub

Private Sub cbxSelDealNum_Change()
  If Me.Dirty Then Call SaveDeal
End Sub

Private Sub cbxSelDealNum_AfterUpdate()
  On Error GoTo cbxSelDealNum_AfterUpdateErr
  With Me
    .Filter = BuildCriteria("lngDealNum", dbGUID, !cbxSelDealNum.Column(0))
    .FilterOn = True
  End With
  Exit Sub
cbxSelDealNum_AfterUpdateErr:
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstThisForm & _
            " : cbxSelDealNum_AfterUpdate"
End Sub

Private Sub cmbSave_Click()
  Call SaveDeal
End Sub

Private Sub cmbNewDeal_Click()
  On Error GoTo cmbNewDeal_ClickErr
  Me.FilterOn = False
  Me.Filter = ""
  Call NewRec(Me)
  Exit Sub

cmbNewDeal_ClickErr:
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstThisForm & " :  cmbNewDeal_Click"
End Sub

Private Sub cmbUndoRec_Click()
  Call UndoRec(Me)
End Sub

Private Sub cmbDelete_Click()
  Call DeleteCurrentRec(Me!lngDealNum)
  Forms(cstMainForm).SetFocus
  Forms(cstMainForm).Requery
End Sub

Private Sub cmbToList_Click()
  If Me.Dirty Then Call SaveDeal
  strCallingForm = cstThisForm
  Forms(cstFilterForm)!tbxCallingForm = strCallingForm
  Me.Visible = False
  Call RefreshActivateMainForm
End Sub

Private Sub Form_Current()
  Dim strFilter As String
  Dim var As Variant
  
  On Error Resume Next
  With Me
    !RecCount = .RecordsetClone.RecordCount
  
    !dteDealDateIn.SetFocus
    !txtContact = ""
     blnNewDeal = .NewRecord
    ' Reset header ctls: deal number selection box & related controls:
    !lblNewDeal.Visible = blnNewDeal
    !lblSelect.Visible = (Not blnNewDeal)
    !cbxSelDealNum.Visible = (Not blnNewDeal)
      
    If Not blnNewDeal Then
      If IsNull(!lngDealNum) Then   'if navigation buttons visible
        MsgBox "Deal# is null: end of current set reached.", vbExclamation + vbOKOnly, _
                cstThisForm & ": Form_Current"
        DoCmd.GoToRecord acActiveDataObject, , acPrevious
        Exit Sub
      Else
        If .RecordsetClone.RecordCount <> 0 Then
          !cbxSelDealNum = !lngDealNum
          !cbxSelDealNum.Requery
          !cbxDealAnalyst.Requery
          !cbxDealAnalyst2.Requery
          !cbxDealAnalystMarket.Requery
          !cbxDealSourceName.Requery
          If Not IsNull(!lngDealIssuerNum) Then !cbxDealIssuer.Requery
          !cbxDealStage.Requery
          !cbxDealDispo.Requery
        End If
      End If
    Else   'new rec
      !cbxSelDealNum = Null
      !cbxDealType = Null
      !cbxDealAnalyst = Null
      !cbxDealAnalyst2 = Null
      !cbxDealAnalystMarket = Null
      !cbxDealSourceName = Null
      !txtContact = ""
      !cbxDealIssuer = Null
      !cbxDealStage = Null
      !cbxDealDispo = !cbxDealDispo.ItemData(0)
    End If
  
    If Err.Number = 3314 Then Err.Clear
    On Error GoTo Form_CurrentErr
    
    If IsLoaded(cstIssuerForm) Then         'synchronize form per current deal
      If Forms!frmIssuer.Visible = True Then
        If IsNull(!lngDealIssuerNum) Then
           strFilter = "[lngIssuerNum] = 1"
        Else
           strFilter = "[lngIssuerNum] = " & !lngDealIssuerNum
        End If
        Forms!frmIssuer.FilterOn = True
        Forms!frmIssuer.Filter = strFilter
        If strCallingForm <> cstDealForm Then .SetFocus
      End If
    End If
    
    If (!cbxDealDispo = "Closed" And IsNull(!cbxSelRole)) Then
      MsgBox "Albion's Role is undetermined even though this deal is closed!", _
              vbExclamation, "Enter missing data"
      !cbxSelRole.SetFocus
    End If
  End With
  
  Exit Sub
    
Form_CurrentErr:
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstThisForm & ": Form_Current"
End Sub

Private Sub Form_Open(Cancel As Integer)
  blnNewSource = False
  blnNewIssuer = False
  strCallingForm = CheckFormArgs(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set cbxCallingBox = Nothing
End Sub

Private Sub txtDateDisp_DblClick(Cancel As Integer)
  Me!txtDateDisp = Date
End Sub

Private Sub lblFilter_Click()
  strCallingForm = cstThisForm
  'Call GetFilteredRecords(cstDealForm)
  'DoCmd.OpenForm cstFilterForm, acNormal
  Call ShowFilterForm(strCallingForm, False)
End Sub

Private Sub lblSourceTab_Click()
  Call OpenThisForm(cstSourceForm, Me!cbxDealSourceName)
End Sub

Private Sub lblIssuerTab_Click()
  Call OpenThisForm(cstIssuerForm, Me!cbxDealIssuer)
End Sub

Private Sub lblSponsorTab_Click()
  Call OpenThisForm(cstSourceForm, Me!cbxEqSponsor)
End Sub

Private Sub lblStatsTab_Click()
  On Error GoTo lblStatsTab_ClickErr
  If Not CheckSourceFields(Me!cbxDealIssuer) Then Exit Sub
  If Me.NewRecord Then
    Call ShowStatusForm(cstThisForm, True)
  Else
    Call ShowStatusForm(cstThisForm)
  End If
  Exit Sub
  
lblStatsTab_ClickErr:
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstThisForm & ": lblStatsTab_Click"
End Sub

Private Sub txtSourceName_DblClick(Cancel As Integer)
  Call lblSourceTab_Click
End Sub
'----------------------------------------------------------------------

Sub CheckSecAmts()
  Dim sfrmSec As Form
  Dim ctl As Control
  Dim strCtl As String
  Dim i As Integer
  
  Set sfrmSec = Me!ctlDetailSubform.Form
  For i = 1 To 4
    strCtl = Choose(i, "sglTrancheSize", "sglAmtOffered", "sglSecInvestSize", "sglEqCoInvest")
    Set ctl = sfrmSec.Controls(strCtl)
    ctl.Value = BillionAmtCheck(ctl, cstMillionMultiple)
    Set ctl = Nothing
  Next i
  Set sfrmSec = Nothing
End Sub

Function CheckSourceFields(cbxCaller As ComboBox) As Boolean
  Dim strMsg As String
  Const cstDefMsg = "Empty Issuer, Source or Sponsor name: No details to view."
  
  CheckSourceFields = True
  
  If cbxCaller.Tag <> 3 Then
    strMsg = strMsg & Choose(cbxCaller.Tag, "The Deal Source ", "The Deal Issuer ")
    strMsg = strMsg & "Name is missing. Please enter."
  End If

  If IsNull(cbxCaller) Then
    CheckSourceFields = False
    If cbxCaller.Tag <> 3 Then
      strMsg = strMsg & vbCrLf & cstDefMsg
      MsgBox strMsg, vbInformation, "Source or Issuer Field Not Set"
    Else
      MsgBox cstDefMsg, vbInformation, "Sponsor Field Not Set"
    End If
  End If
  cbxCaller.SetFocus
End Function

Sub OpenThisForm(strForm As String, cbxCaller As ComboBox)
  Dim strLink As String
  On Error GoTo OpenThisFormErr
  
  strCallingForm = cstThisForm
  If Not CheckSourceFields(cbxCaller) Then Exit Sub
  
  If strForm = cstIssuerForm Then
    strLink = "[lngIssuerNum]= " & cbxCaller.Column(0)
  Else
    strLink = "[lngSourceNum]= " & cbxCaller.Column(0)
  End If
  DoCmd.OpenForm strForm, , , strLink, , , cstThisForm
  Exit Sub
  
OpenThisFormErr:
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstThisForm & ": OpenThisForm"
End Sub

Function SaveDeal() As Boolean
  Dim lngRec As Long
  On Error GoTo SaveDealErr
  
  Call CheckSecAmts
  SaveDeal = True
  lngRec = SaveRec(Me)
  Debug.Print "SaveDeal blnRequeryCbx: " & blnRequeryCbx
  If blnRequeryCbx Then
    Debug.Print "Me!cbxSelDealNum.ItemData(0): " & Me!cbxSelDealNum.ItemData(0)
    Me!cbxSelDealNum.Requery
    Me!cbxSelDealNum = Me!cbxSelDealNum.ItemData(0)
     
    'Call cbxSelDealNum_AfterUpdate
    blnRequeryCbx = False
  'Else
  '  If lngRec <> 0 Then DoCmd.GoToRecord acActiveDataObject, , acFirst
  End If
  Exit Function
  
SaveDealErr:
  SaveDeal = False
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstThisForm & ": SaveDeal"
End Function

Function ChangeOfAnalyst(cbx As ComboBox) As Boolean
' Possible cbx's: cbxDealAnalyst; cbxDealAnalyst2, cbxDealAnalystMarket
  On Error GoTo ChangeOfAnalystErr
  'If IsNull(cbx) Then Exit Function
  'If IsNull(cbx.OldValue) Then Exit Function
  If cbx = cbx.OldValue Then Exit Function
  ChangeOfAnalyst = Not cbx.Column(2) 'second visible col=current employee flag (y/n)
  If ChangeOfAnalyst Then
    Beep
    MsgBox "You can assign deals to current employees only.", vbInformation
  End If
  Exit Function
  
ChangeOfAnalystErr:
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstThisForm & ": ChangeOfAnalyst"
End Function

Sub DealEntitiesAfterUpd(cbx As ComboBox)
  Dim iResponse As Integer, iTag As Integer
  Dim strData As String, strMsg As String, strTitle As String, strDef As String
 ' On Error GoTo DealEntitiesAfterUpdErr
  
'  Set cbxCallingBox = cbx 'set global var
  
  iTag = cbx.Tag
  strMsg = "Enter the current "
  strMsg = strMsg & Choose(iTag, "Source Name:", "Issuer Name:", "Sponsor Name:")
  strTitle = Choose(iTag, "New Source Entry", "New Issuer Entry", "New Sponsor Entry")
  strDef = Choose(iTag, "<source name>", "<issuer name>", "<sponsor name>")
  If cbx = -1 Then
    strData = InputBox(strMsg, strTitle, strDef)
    If DealEntityNotInList(cbx, strData) = 0 Then Exit Sub
  End If
  If blnNewSource Then
    cbx.Requery
    If iTag = 1 Then  'req other cbx with sources as input
      Me!cbxEqSponsor.Requery
    Else
      Me!cbxDealSourceName.Requery
    End If
  End If
  If iTag = 2 Then
    If blnNewIssuer Then cbx.Requery
    Me!lblStatsTab.Visible = True
  End If
  Exit Sub
  
DealEntitiesAfterUpdErr:
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstThisForm & ": DealEntitiesAfterUpd"
End Sub

Function DealEntityNotInList(cbx As ComboBox, strNewData As String) As Integer
'Returns the Response expected by the NotInList event (if LimitToList=Yes).
' strOrigin refer to either 'Source' or 'Equity Sponsor'.
  Dim rst As DAO.Recordset
  Dim frm As Form
  Dim iTag As Integer
  Dim strMsg As String, strTitle As String, strOrigin As String, strForm As String
  Dim strRespLast As String, strRespFirst As String, strIssuer As String
  Dim blnProceed As Boolean
  
  blnNewSource = False: blnNewIssuer = False  'reset global var
  'On Error GoTo DealEntityNotInListErr
  
  If Len(Trim(strNewData)) = 0 Then
    'Ask?: do you want to delete cbx.olddata from the main list?
    DealEntityNotInList = acDataErrContinue '=0
    cbx.Undo
    Exit Function
  End If
  
  iTag = cbx.Tag
  strForm = Choose(iTag, cstSourceForm, cstIssuerForm, cstSourceForm)
  strOrigin = Choose(iTag, "Source", "Issuer", "Sponsor")
  strMsg = strNewData & " is not on file. " & vbCrLf & "Is this a new " & strOrigin & "?"
  strTitle = "Unknown " & strOrigin
  
  If MsgBox(strMsg, vbYesNo + vbQuestion, strTitle) = vbNo Then
    DealEntityNotInList = acDataErrContinue '=0
    cbx.Undo
  Else
  
InputStart:
    If iTag = 1 Or iTag = 3 Then
      strRespLast = Trim(InputBox("Please, enter the contact's last name:", _
                                  "New Contact", "<last name>"))
      If strRespLast = "" Then
        If MsgBox("The last name is required." & vbCrLf & _
                   "Do you want to cancel this new source entry?", vbExclamation + vbYesNo, _
                   "Missing Last Name") = vbYes Then
          Exit Function
        Else
          GoTo InputStart 'try again
        End If
      End If
      strRespFirst = Trim(InputBox("Please, enter the contact's first name:", _
                                 "New Contact", "<first name>"))
    Else
      strIssuer = Trim(InputBox("Please, enter the Issuer's full name:", _
                                  "New Issuer", "<full name>"))
      If strIssuer = "" Then
        If MsgBox("The Issuer's name is required." & vbCrLf & _
                   "Do you want to cancel this new Issuer entry?", vbExclamation + vbYesNo, _
                   "Missing Issuer Name") = vbYes Then
          Exit Function
        Else
          GoTo InputStart 'try again
        End If
      End If
    End If

    'Ref form w/o showing it:
    If Not IsLoaded(strForm) Then DoCmd.OpenForm strForm, acNormal, , , , , cstDealForm
    Set frm = Forms(strForm)
    Set rst = frm.RecordsetClone
    With rst
      .AddNew
      If iTag = 1 Or iTag = 3 Then
        ' Add new rec in rst:
        blnNewSource = True
        !txtSourceName = strNewData
        !txtSourceContactLast = strRespLast
        !txtSourceContactFirst = strRespFirst
        !lngSourceTypeNum = 9 '  sponsor group
      Else
        blnNewIssuer = True
        !txtIssuerName = strIssuer
      End If
      .Update
      .Bookmark = .LastModified
    End With
    
    frm.Bookmark = rst.Bookmark
    DealEntityNotInList = acDataErrAdded  '=2
       
    If iTag <> 2 Then
      'If blnNewSource Then
      cbx = rst!lngSourceNum
    Else
      'If blnNewIssuer Then
      cbx = rst!lngIssuerNum
    End If
    frm.Visible = True
  
  End If 'Not in list add?=yes
  
  Set rst = Nothing: Set frm = Nothing
  Exit Function
  
DealEntityNotInListErr:
  Set rst = Nothing: Set frm = Nothing
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, cstThisForm & ": DealEntityNotInList"
End Function
