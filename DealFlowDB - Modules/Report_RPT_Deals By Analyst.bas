VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_RPT_Deals By Analyst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
'
' RPT Deals By Analyst May-9-02 16:10
'
'================================================================================
  
Private Sub Report_Close()
  If IsLoaded(cstMainForm) Then
    Forms(cstMainForm).Visible = True
  Else
    DoCmd.OpenForm cstMainForm
  End If
End Sub
