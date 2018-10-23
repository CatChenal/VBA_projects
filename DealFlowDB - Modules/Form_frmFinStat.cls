VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFinStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
' DBFrontEnd
' CLSMDL frmFinStat Jun-14-02 10:30
'
'================================================================================
Const cstThisFrm = cstFinStatusForm

Private Sub cmbCalc_Click()
  Call OpenCalculator
End Sub

Private Sub cmbClose_Click()
  If IsLoaded(cstDealForm) Then
    If Forms!frmDeal.Visible = True Then
      strCallingForm = cstDealForm
      Forms(cstDealForm).SetFocus
    End If
  End If
    
 Call SaveRec(Me)
 DoCmd.Close
End Sub

Private Sub cmbNewStats_Click()
  DoCmd.RunCommand acCmdRecordsGoToNew
  Me.Modal = True
End Sub

Private Sub Form_Open(Cancel As Integer)
  strCallingForm = CheckFormArgs(Me)
  Me!dteFinPeriodEnd.SetFocus
  Call lngFinPerNum_AfterUpdate
End Sub

Private Sub lngFinPerNum_AfterUpdate()
  If Not IsNull(Me!lngFinPerNum) Then
    Me!lblPeriod.Caption = Me!lngFinPerNum.Column(2)
  Else
    Me!lblPeriod.Caption = "."
  End If
End Sub

Private Sub Form_Current()
  Call sglFinEBITDA_AfterUpdate
  If Not IsLoaded(cstDealForm) Then Me!cmbNewStats.Enabled = True
End Sub

Private Sub sglFinEBITDA_AfterUpdate()  'Recalc related ratios
  Me!sglFinEBITDA = BillionAmtCheck(Me!sglFinEBITDA, cstMillionMultiple)
'  Me!sglFinEBITDA = CheckAmtEntry(Me!sglFinEBITDA, cstMillionMultiple)
  Call sglFinIntExp_AfterUpdate     ' interest coverage
  Call sglFinTotDebt_AfterUpdate      ' leverage
End Sub

Private Sub sglFinEquity_AfterUpdate()
  Me!sglFinEquity = BillionAmtCheck(Me!sglFinEquity, cstMillionMultiple)
'  Me!sglFinEquity = CheckAmtEntry(Me!sglFinEquity, cstMillionMultiple)
End Sub

Public Sub sglFinIntExp_AfterUpdate()
  Me!sglFinIntExp = BillionAmtCheck(Me!sglFinIntExp, cstMillionMultiple)
'  Me!sglFinIntExp = CheckAmtEntry(Me!sglFinIntExp, cstMillionMultiple)
  If ((IsNull(Me!sglFinIntExp) Or Me!sglFinIntExp = 0)) Then
    Me!sglIntCove = 0
  Else
    Me!sglIntCove = Me!sglFinEBITDA / Me!sglFinIntExp
  End If
End Sub

Private Sub sglFinSales_AfterUpdate()
 Me!sglFinSales = BillionAmtCheck(Me!sglFinSales, cstMillionMultiple)
 ' Me!sglFinSales = CheckAmtEntry(Me!sglFinSales, cstMillionMultiple)
End Sub

Private Sub sglFinTotDebt_AfterUpdate()
  Me!sglFinTotDebt = BillionAmtCheck(Me!sglFinTotDebt, cstMillionMultiple)
  'Me!sglFinTotDebt = CheckAmtEntry(Me!sglFinTotDebt, cstMillionMultiple)
  If ((IsNull(Me!sglFinEBITDA) Or Me!sglFinEBITDA = 0)) Then
    Me!sglLeverage = 0
  Else
    Me!sglLeverage = Me!sglFinTotDebt / Me!sglFinEBITDA
  End If
End Sub

Private Sub sglFinTotDebt_Cap_AfterUpdate()
'  Me!sglFinTotDebt_Cap = CheckAmtEntry(Me!sglFinTotDebt_Cap, cstMillionMultiple)
  Me!sglFinTotDebt_Cap = BillionAmtCheck(Me!sglFinTotDebt_Cap, cstMillionMultiple)
End Sub
