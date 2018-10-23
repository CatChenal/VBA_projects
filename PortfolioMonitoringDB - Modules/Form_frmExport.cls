VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'
'================================================================================
'
' CLSMDL frmExport Nov-18-02
' May-30-02 21:30
'
'================================================================================
'
Private Sub Form_Load()
  With Me
    .InsideWidth = 4600
    .InsideHeight = 2990
  End With
End Sub

Private Sub lblCancel_Click()
  DoCmd.Close acForm, "frmExport"
End Sub

Private Sub lblExport_Click()
  Call ExportForecastPrep(Me)
End Sub

Private Sub opgExportSel_AfterUpdate()
  Me!cbxSelFrcType.Enabled = (Me!opgExportSel < 2)
End Sub
