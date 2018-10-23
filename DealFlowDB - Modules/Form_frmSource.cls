VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
' DBFrontEnd
' CLSMDL frmSource May-20-02 15:15
'
'================================================================================
Const cstThisForm = "frmSource"
Dim varPrevVal As Variant

Private Sub cbxSourceName_NotInList(NewData As String, Response As Integer)
  MsgBox "Use the command button to enter a new Source.", vbInformation, "Source Not In List"
  Response = acDataErrContinue
  Me!cbxSourceName.Undo
End Sub

Private Sub cmbCancel_Click()
'  If strCallingForm = "" Then
 '   DoCmd.Close
 ' Else
 '   Forms(strCallingForm).SetFocus
  If Me.Dirty Then Me.Undo
  Me.Visible = False
  DoCmd.Close
 ' End If
End Sub

Private Sub cmbClose_Click()
  Call CloseFrm(Me) 'includes saverec
End Sub

Private Sub cmbContact_Click()
  On Error GoTo Err_cmbContact_Click
  DoCmd.Hourglass True
  With Me
    If Not IsNull(!lngSourceNum) Then
      If Not IsNull(!txtSourceContactLast) Then
        Call DisplayAAContact(!txtSourceContactLast, !txtSourceContactFirst, !cbxSourceName)
      Else
        MsgBox "Searching the Outlook Contacts folder for this" & vbCrLf & _
               "contact is not possible without the last name.", vbExclamation, "Cannot Search for Contact"
      End If
    End If
 
    'Update changes if necessary:
    If Not IsNull(!blnSourceOLFlag) Then
      If !blnSourceOLFlag <> (blnAddedContact Or blnContact) Then
        'update flag field in table:
        !blnSourceOLFlag = (blnAddedContact Or blnContact)
      End If
    Else  'update without condition
      !blnSourceOLFlag = (blnAddedContact Or blnContact)
    End If
  End With
  
Exit_cmbContact_Click:
  DoCmd.Hourglass False
  Exit Sub

Err_cmbContact_Click:
  MsgBox Err & ": " & Err.Description, , cstThisForm & " cmbContact_Click"
  Resume Exit_cmbContact_Click
End Sub

Private Sub cmbNewContact_Click()
  Dim txtCo As String
  Dim lngType As Long
  
  If MsgBox("Enter a new contact for this company?", vbQuestion + vbOKCancel) = vbOK Then
    With Me
      txtCo = !cbxSourceName
      lngType = !cbxSourceTypeNum
      Call cmbNewSource_Click
      !cbxSourceName = txtCo
      !cbxSourceTypeNum = lngType
      !txtSourceContactLast.SetFocus
    End With
  End If
End Sub

Private Sub cmbNewSource_Click()
  Call NewRec(Me)
End Sub

Private Sub Form_AfterInsert()
  If IsLoaded(cstDealForm) Then
    If Forms!frmDeal.Visible = True Then
      If strCallingForm = cstDealForm Then
        Forms!frmDeal.SetFocus
        cbxCallingBox.Requery
      End If
    End If
  End If
End Sub

Private Sub Form_AfterUpdate()
  Me.Modal = False
End Sub

Private Sub Form_Current()
  'varPrevVal = Null
  If Me.NewRecord = True Then blnNewSource = True
  'Call CheckContact
End Sub

Private Sub Form_Load()
  varPrevVal = Null
  strCallingForm = CheckFormArgs(Me)
End Sub

Private Sub txtSourceContactFirst_BeforeUpdate(Cancel As Integer)
  varPrevVal = Me!txtSourceContactFirst.OldValue
End Sub

Private Sub txtSourceContactFirst_AfterUpdate()
 Dim frm As Form
  
  If varPrevVal <> Me!txtSourceContactFirst Then  'value changed
    Call CheckContact
    If IsSet(cbxCallingBox) Then
      Set frm = cbxCallingBox.Parent
      frm.SetFocus
      cbxCallingBox = Me!lngSourceNum
      cbxCallingBox.Requery
      Me.SetFocus
      Set frm = Nothing
    End If
  End If
End Sub

Private Sub txtSourceContactLast_BeforeUpdate(Cancel As Integer)
  varPrevVal = Me!txtSourceContactLast.OldValue
End Sub

Private Sub txtSourceContactLast_AfterUpdate()
  Dim frm As Form
  
  If (varPrevVal <> Me!txtSourceContactLast) Then 'value changed
    Call CheckContact
    If IsSet(cbxCallingBox) Then
      Set frm = cbxCallingBox.Parent
      frm.SetFocus
      cbxCallingBox.Requery
      frm.Recalc
      Me.SetFocus
      Set frm = Nothing
    End If
  End If
End Sub

Private Sub cbxSourceName_AfterUpdate()
  Dim frm As Form
  Dim strMsg As String
  On Error GoTo cbxSourceName_AfterUpdateErr

  If Not IsNull(varPrevVal) Then
    If varPrevVal <> Me!cbxSourceName Then
      strMsg = "You have modified the Source name of the current Deal." & vbCrLf & _
               "Replace '" & varPrevVal & "' with '" & Me!cbxSourceName & "'?"
      If MsgBox(strMsg, vbQuestion + vbYesNo, "Deal Source Change") = vbNo Then
        Me!cbxSourceName = varPrevVal
        Me.Dirty = False
      Else
        If IsSet(cbxCallingBox) Then
            Set frm = cbxCallingBox.Parent
            frm.SetFocus
            cbxCallingBox.SetFocus
            cbxCallingBox.Requery
            'cbxCallingBox = Me!cbxSourceName
            Set frm = Nothing
        End If
        Call CheckContact
      End If
    End If
  End If

cbxSourceName_AfterUpdateExit:
  Set frm = Nothing
  Exit Sub

cbxSourceName_AfterUpdateErr:
  MsgBox Err & ": " & Err.Description, , cstThisForm & " cbxSourceName_AfterUpdate"
  Resume cbxSourceName_AfterUpdateExit
End Sub

Private Sub cbxSourceName_BeforeUpdate(Cancel As Integer)
  varPrevVal = Me!cbxSourceName.OldValue
  'If CStr(varPrevVal) <> Me!cbxSourceName Then Cancel = CheckDuplicateName
  'If Cancel = 0 Then
  Cancel = CheckDirectCo
End Sub

Private Sub CheckContact()
' To give a visual clue as to the value of the 'OLFlag' (Is this contact in Outlook?)
' and entice the user to click on it to perform a search and/or a new contact entry in OL and
' update the current record with the outcome.
  Dim blnOLFlag As Boolean
  With Me
    If IsNull(!lngSourceNum) Then
      !cmbContact.Visible = False
    Else
      If Not IsNull(!blnSourceOLFlag) Then blnOLFlag = !blnSourceOLFlag
      Call ShowContactButton(blnOLFlag)     'show/hide contact icon as per source contact data
    End If
  End With
End Sub

Public Sub ShowContactButton(bln As Boolean)
  With Me
    !cmbContact.Visible = True
    If bln Then
      !cmbContact.PictureData = !cmbContactY.PictureData
    Else
      !cmbContact.PictureData = !cmbContactN.PictureData
    End If
  End With
End Sub

Function CheckDirectCo() As Integer
  'On Error GoTo CheckDirectCoErr
  
  If Not Me.Dirty Then Exit Function
  ' Type = Direct from co?
  If Me!cbxSourceTypeNum = 1 And Me!cbxSourceName <> "Direct" Then
    CheckDirectCo = 9
    If MsgBox("If " & UCase(Me!cbxSourceName) & _
              " is also the Issuer, then leave the Source name as 'Direct'" & vbCrLf & _
              "Do you want to undo this record?", _
                  vbExclamation + vbYesNo, "Direct Source Check") = vbYes Then Me.Undo
  End If
'  Exit Function

'CheckDirectCoErr:
'  MsgBox Err & ": " & Err.Description, , cstThisForm & " CheckDirectCo"
End Function

Function CheckDuplicateName() As Integer
  Dim dbs As DAO.Database
  Dim rst As DAO.Recordset
  Dim strSQL As String, strCrit As String
  On Error GoTo CheckDuplicateNameErr
  
  If Not Me.Dirty Then Exit Function
  If Me!cbxSourceName = "Direct" Then Exit Function
  
  Set dbs = CurrentDb
  strSQL = "SELECT * from tblSource where txtSourceName='" & Me!cbxSourceName & "'"
  Set rst = dbs.OpenRecordset(strSQL)
  If rst.AbsolutePosition <> -1 Then
    rst.MoveLast
    rst.MoveFirst
    strCrit = "[txtSourceContactLast] = """ & Me!txtSourceContactLast & """ AND "
    strCrit = strCrit & "[txtSourceContactFirst] = """ & Me!txtSourceContactFirst & """"
    rst.FindFirst strCrit
    If Not rst.NoMatch Then      'cannot add: already there
       CheckDuplicateName = 9
       If MsgBox("This Source Contact already exists!" & vbCrLf & "Do you want to undo this record?", _
                  vbExclamation + vbYesNo, "Duplicate Check") = vbYes Then Me.Undo
        
      GoTo CheckDuplicateNameExit
    End If
  End If
  
CheckDuplicateNameExit:
  rst.Close
  Set rst = Nothing
  Set dbs = Nothing
  Exit Function

CheckDuplicateNameErr:
  MsgBox Err & ": " & Err.Description, , cstThisForm & " CheckDuplicateName"
  Resume CheckDuplicateNameExit
End Function
