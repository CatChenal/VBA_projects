VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_SRPT_FinStatPerDeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
'
' RPT DATA FinStatPerDeal May-6-02 10:10
'
'================================================================================

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
  If intFinProForma = 0 Then
    lblPF.Caption = "Pro Form"
  Else
    lblPF.Caption = "Actual"
  End If
End Sub
