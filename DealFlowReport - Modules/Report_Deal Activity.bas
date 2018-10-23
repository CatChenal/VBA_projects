VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Deal Activity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub GroupHeader4_Print(Cancel As Integer, PrintCount As Integer)
  If IsNull(txtSecType) And IsNull(sglTrancheSize) And IsNull(sglAmtOffered) Then
    Cancel = 1
  Else
    Cancel = 0
  End If
End Sub
