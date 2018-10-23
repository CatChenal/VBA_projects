VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'===============================================================================
' DBFrontEnd
' CLSMDL frmSecurity May-10-02 15:06
'
'===============================================================================
Const cstFrm = "frmSecurity"
Const cstFrmHeight = 1440 * 2.5
Const cstFullHeight = 1440 * 3.5  '4.8  ' 1440 twips = 1 inch.

Private Sub Form_Current()
  With Me
    If !lngSecTypeNum = 7 Then 'Common Stock
      !sglSecFee.Enabled = False
    Else
      !sglSecFee.Enabled = True
    End If
  End With
End Sub

Private Sub Form_Open(Cancel As Integer)
  Dim str As String
  
  On Error Resume Next
  str = Parent.Name
  If Err = 2452 Then 'no parent: form opened on its own
    Me!lblNewSec.Visible = False
    Me!lblDelSec.Visible = False
    strCallingForm = cstMainForm
  Else
    If Parent.Name = cstDealForm Then
      Me!lblNewSec.Visible = True
      Me!lblDelSec.Visible = True
      strCallingForm = cstDealForm
    End If
  End If
  
End Sub

Private Sub lblDelSec_Click()
  On Error Resume Next
  DoCmd.RunCommand acCmdDeleteRecord
  If Err.Number = 2046 Then 'nothing to delete
    MsgBox "There is no record to delete", vbInformation
    Err.Clear
    Exit Sub
  End If
End Sub

Private Sub lblNewSec_Click()
  Dim lngDeal As Long
  
  lngDeal = Parent!lngDealNum
  DoCmd.GoToRecord , , acNewRec
  Me!lngSecDealNum = lngDeal
End Sub
