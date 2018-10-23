VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_RPT_Deals List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
'
' RPT Deals List May-16-02 14:40
'
'================================================================================
  
Private Sub Report_Close()
  If IsLoaded(cstMainForm) Then
    Forms(cstMainForm).Visible = True
  Else
    DoCmd.OpenForm cstMainForm
  End If
End Sub

Private Sub Report_NoData(Cancel As Integer)
  MsgBox "No Data To Preview", vbExclamation, "Operation Cancelled"
  Cancel = True
End Sub
