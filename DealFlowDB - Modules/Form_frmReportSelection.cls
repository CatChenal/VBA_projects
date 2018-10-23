VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmReportSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'===============================================================================
' DBFrontEnd
' CLSMDL frmReportSelection Sep-05-03 14:40
' Changes: simplified cmbPreview_Click
' Prev: Jan-30-03 15:20
'===============================================================================
Const cstThisForm = cstReportForm

Private Sub cmbClose_Click()
  Me.Visible = False
  Call ApplyFormFilter(cstMainForm)
End Sub

Private Sub cmbPreview_Click()
  Dim strDoc As String
  Dim blnNoSelection As Boolean, blnFormRedrawn As Boolean
  Dim i As Integer, iDocType As Integer
  '---------------------------------------------------------------------------
  On Error GoTo cmbPreviewErr
  iDocType = Me!opgSelType
  
  blnNoSelection = (Me!lbxReportList.ItemsSelected.Count = 0)
  If blnNoSelection Then
    If iDocType = 1 Then
      strDoc = "No report selected."
    Else
      strDoc = "No query selected."
    End If
    MsgBox strDoc, vbExclamation + vbOKOnly, "Selection Missing"
    Exit Sub
  Else
    i = Me!lbxReportList.ListIndex
    strDoc = Me!lbxReportList.ItemData(i)
  End If
  
  Select Case iDocType
    Case 1  ' ----------------------------1: REPORTS
        DoCmd.OpenReport strDoc, acViewPreview
      
    Case 2  ' ----------------------------2: QUERIES
      DoCmd.Hourglass True
      strDoc = Me!lbxReportList.ItemData(i)
      ' Check if subform data source has to change:
      '  Do not set source object to query result form until new subform is ready: (frmDatasheetSub)

      ' Change from hidden to icon
      DoCmd.OpenForm cstDataForm, , , , , acIcon

      blnFormRedrawn = (Len(Forms(cstDataForm)!frmDatasheetSub.SourceObject) = 0)
      If Not blnFormRedrawn Then
        blnFormRedrawn = (Forms(cstDataForm)!frmDatasheetSub.Form.RecordSource <> strDoc)
      End If
                    
      If blnFormRedrawn Then   'recreate subform
        Forms(cstDataForm).Visible = False
        Forms(cstMainForm).Visible = True
        DoCmd.Close acForm, cstDataForm
        ' Reset bln flag with redrawing function
        blnFormRedrawn = RedrawForm(cstDataSubForm, strDoc, iDocType)
        If Not blnFormRedrawn Then
          MsgBox "Contact Your DB Admin", vbExclamation, "Error Redrawing Display Form"
          GoTo cmbPreviewExit
        End If
        DoCmd.OpenForm cstDataForm, , , , , acHidden
      End If
  
      Forms(cstDataForm)!frmDatasheetSub.SourceObject = cstDataSubForm '"frmDatasheetSub"
      Forms(cstDataForm)!frmDatasheetSub.Form.RecordSource = strDoc
   
      With Forms(cstDataForm)
        If !frmDatasheetSub.Form.RecordsetClone.RecordCount = 0 Then
          DoCmd.Hourglass False
          MsgBox "No data for your selection.", vbExclamation + vbOKOnly, "Query returned no records"
          Exit Sub
        Else
          .Visible = True
          DoCmd.Maximize
          !txtQryName = Me!lbxReportList.Column(1, i)  'caption
        End If
      End With
    
  End Select
    
cmbPreviewExit:
  Me.Visible = False
  DoCmd.Hourglass False
  Exit Sub
   
cmbPreviewErr:
  DoCmd.Hourglass False
  MsgBox Err.Number & ": " & Err.Description, , cstThisForm & ": cmbPreview"
End Sub

Private Sub Form_Activate()
  Call ApplyFormFilter(cstThisForm)  'includes ResetFilterClues
End Sub

Private Sub Form_Load()
  Me!opgSelType = 1
End Sub

Private Sub lblFilter_Click()
  Call ShowFilterForm(cstThisForm, False)
End Sub

Private Sub lbxReportList_AfterUpdate()
  Dim i As Integer
  With Me
    i = !lbxReportList.ListIndex
    !txtRptDesc = !lbxReportList.Column(2, i)
  End With
End Sub

Private Sub lbxReportList_DblClick(Cancel As Integer)
  Call cmbPreview_Click
End Sub

Private Sub opgSelType_AfterUpdate()
  Me!lbxReportList.Requery
End Sub
