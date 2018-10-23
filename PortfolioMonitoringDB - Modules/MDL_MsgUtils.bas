Attribute VB_Name = "MDL_MsgUtils"
Option Compare Database
Option Explicit
'================================================================================
'
' MDL MsgUtils Apr-4-03 10:35
'
'================================================================================

Function SetMsgForm()
  Dim strOpenArgs As String
  Const cstDefHeightWidth = 2.9 * 1440
  strOpenArgs = ""
  With Forms(cstFRM_Msg)
    .InsideWidth = cstDefHeightWidth
    .InsideHeight = cstDefHeightWidth
    If Len(.OpenArgs & "") > 0 Then
      strOpenArgs = .OpenArgs
      .AllowEdits = True
      !txtMsg = " "
      !txtMsg = strOpenArgs
      !txtMsg.SetFocus
      !txtMsg.SelStart = 0
      .AllowEdits = False
    End If
  End With
End Function
