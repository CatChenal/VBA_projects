Attribute VB_Name = "MDL ReCreateSubForm"
Option Compare Database
Option Explicit

'================================================================================
' DBFrontEnd
' MDL ReCreateSubForm Jan-10-03 15:40
'
'================================================================================
'
Public Function RedrawForm(strTemplateForm As String, strRecSource As String, _
                           Optional intTblOrQry As Integer = 1) As Integer
' Usage: To change the recordset(fields) of a form displayed in datasheet view.
' Returns 0(False) or -1(True)
'?RedrawForm("frmDatasheetSub","qryYearAllDealsStats", 2) ' "qryYearAllDealsTotals"
'?RedrawForm("frmDatasheetSub","tlkpFund", 1)
    Dim frm As Form
    Dim ctl As Control
    Dim tbx As TextBox
    Dim fld As DAO.Field
    Dim fldsColl As DAO.Fields
    Dim strFld As String, strCtl As String
    Dim iTop As Integer 'in twips, with whole numbers, no need to be a Long
    Dim c As Integer, iCtlCount As Integer
    Dim blnResult As Boolean
    c = 0
    On Error GoTo ProcErr
    '---------------------------------------------
    Set dbs = CurrentDb
    DoCmd.OpenForm strTemplateForm, acDesign, , , , acHidden
    Set frm = Forms(strTemplateForm)
    ' Delete previous controls & change recsource:
    c = frm.Controls.Count
    Do While c > 0
      Set ctl = frm.Controls(c - 1)
      strCtl = ctl.Name
      Application.DeleteControl frm.Name, strCtl
      c = frm.Controls.Count
      If IsEmpty(frm.Controls) Then Exit Do
    Loop
    frm.RecordSource = strRecSource
    
    If Not IsMissing(intTblOrQry) And intTblOrQry = 2 Then 'qry datasource
      Set fldsColl = dbs.QueryDefs(strRecSource).Fields
    Else
      Set fldsColl = dbs.TableDefs(strRecSource).Fields
    End If
    
    iTop = 100
    For Each fld In fldsColl
      strFld = fld.Name
      Set tbx = CreateControl(frm.Name, acTextBox, acDetail, , , 1600, iTop, 1000, 220)
      tbx.Name = strFld
      tbx.ControlSource = strFld
      iTop = iTop + 300
    Next fld
    DoCmd.Close acForm, strTemplateForm, acSaveYes
    blnResult = True
    
ProcExit:
  Set ctl = Nothing
  Set tbx = Nothing
  Set frm = Nothing
  Set fldsColl = Nothing
  Set dbs = Nothing
  RedrawForm = blnResult
  Exit Function
  
ProcErr:
  blnResult = False
  MsgBox "Error: " & Err & "; " & Err.Description, vbExclamation, "RedrawForm"
  Resume ProcExit
End Function
