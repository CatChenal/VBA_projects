VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_RPT_Current Deal Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
'
' RPT Current Deal Data May-9-02 16:10
'
'================================================================================

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
  If Len(txtDealMemo & "") = 0 Then Report.Section(acDetail).Visible = False
End Sub
