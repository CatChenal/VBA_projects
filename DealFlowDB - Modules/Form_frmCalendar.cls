VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
'
' CLSMDL frmCalendar Apr-26-02 12:00
'
'================================================================================

Private Sub cmbClose_Click()
  DoCmd.Close
End Sub

Private Sub cmdOK_Click()
  Call UpdateDateFromCal(False)
  With Me
    !txtForm = ""
    !txtCtl = ""
    !PreVal = ""
  End With
  DoCmd.Close
End Sub

Private Sub Form_Open(Cancel As Integer)
  With Me
    !acxCal.Value = .OpenArgs
    !PreVal = .OpenArgs
  End With
End Sub
