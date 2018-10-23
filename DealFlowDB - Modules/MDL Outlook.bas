Attribute VB_Name = "MDL Outlook"
Option Compare Database
Option Explicit
'================================================================================
' DBFrontEnd
' MDL Outlook: Aug-29-03
'     Prev: Apr-29-02 17:05
'
'================================================================================
Public blnAddedContact As Boolean     'set in DisplayContact
Public blnContact As Boolean          'set in DisplayContact
Dim ol As Outlook.Application
Dim ns As Outlook.NameSpace

Const cstPublicFoldersDir = "Public Folders"
Const cstAllPublicFoldersDir = "All Public Folders"
Const cstAAPublicSubfolderDir = "Albion Alliance"
Const cstAAPublicSubfolder = "Albion Alliance Contacts"
Const cstAAContactsForm = "IPM.Contact"

Const cstMDL = "MDL Outlook"  'module contains functions to interface with MS Outlook

'------------------------------------------------------------------------------------------

Function StartOutlook()
  On Error GoTo StartOutlookErr
  Set ol = New Outlook.Application
  Set ns = ol.GetNamespace("MAPI")                  'Reference the MAPI layer
  ns.Logon "MS Exchange Settings", , False, True    'New session with default settings, no dialog
  
StartOutlookExit:
  Exit Function
  
StartOutlookErr:
  Set ol = Nothing
  Set ns = Nothing
  MsgBox "Error (" & Err.Number & "): " & Err.Description, vbExclamation, "StartOutlook"
  Resume StartOutlookExit
End Function
  
Function EndOutlook()
  If Not IsSet(ns) Then
    Set ol = Nothing
    Exit Function
  End If
  ns.Logoff         'End current session
  Set ns = Nothing  'Release memory
  Set ol = Nothing
End Function

Public Function GetAAContactsFolder(mapiAllFolders As Outlook.MAPIFolder) As Outlook.MAPIFolder
' Level added since migration to AC domain.
  Set GetAAContactsFolder = mapiAllFolders.Folders(cstAAPublicSubfolderDir).Folders(cstAAPublicSubfolder)
End Function

Public Function DisplayAAContact(strLast As String, strFirst As String, strCo As String)
  Dim fdrAAContacts As Outlook.MAPIFolder
  Dim itmContact As Outlook.ContactItem
  Dim frmContact As Outlook.Inspector
  Dim strCriteria As String
  
  blnContact = False
  On Error GoTo DisplayAAContactErr

  Call StartOutlook                       'sets ol (app), ns (namespace)
  Set fdrAAContacts = GetAAContactsFolder(ns.Folders(cstPublicFoldersDir).Folders(cstAllPublicFoldersDir))
  
  'Establish the criteria and locate the Contact
  strCriteria = "[LastName] = '" & strLast & "' AND [FirstName] = '" & strFirst & "'"
  Set itmContact = fdrAAContacts.Items.Find(strCriteria)
  If Not (itmContact Is Nothing) Then blnContact = True
  
  If Not blnContact Then
    If MsgBox("Would you like to enter this contact, " & _
              strFirst & " " & strLast & ", in Outlook?", _
              vbQuestion + vbYesNo, "Contact not found in Outlook") = vbYes Then   'add
      Set itmContact = fdrAAContacts.Items.Add(cstAAContactsForm)
      With itmContact
        .LastName = strLast
        .FirstName = strFirst
        .CompanyName = strCo
        .Save
      End With
      blnAddedContact = True
      Set frmContact = itmContact.GetInspector
      frmContact.Display
    Else
      blnAddedContact = False
    End If
  Else
    With itmContact
      'Make sure correct form is bound to contact item:
      If .MessageClass <> cstAAContactsForm Then
         .MessageClass = cstAAContactsForm
         .Save
      End If
      Set frmContact = .GetInspector  'Display the Contact
      frmContact.Display
    End With
  End If
   
DisplayAAContactExit:
  Call EndOutlook
  Set fdrAAContacts = Nothing
  Set itmContact = Nothing
  Exit Function
  
DisplayAAContactErr:
  MsgBox Err & ": " & Err.Description, vbExclamation, "DisplayAAContact"
  Resume DisplayAAContactExit
End Function

Public Function UpdateOLContact(frmUsed As Form, strLast As String, strFirst As String, _
                                blnCreateIfNotFound As Boolean)
  Dim fdrAAContacts As Outlook.MAPIFolder
  Dim itmContact As Outlook.ContactItem
  Dim strCriteria As String
  Dim blnSave As Boolean
  Const cstProc = "UpdateOLContact"
  
  strCriteria = "": blnSave = False
  
  On Error GoTo UpdateOLContactErr
    
  If frmUsed.Name <> cstMgmtForm Then
    MsgBox "Function designed to work with the form 'frmManagement' only.", _
           vbExclamation, cstMDL & cstProc & ": wrong use"
    Exit Function
  End If
  If (Len(strLast & "") = 0 Or Len(strFirst & "") = 0) Then
    MsgBox "Both First & Last names are required.", vbExclamation, _
           cstMDL & cstProc & ": data check"
    Exit Function
  End If

  'Establish the criteria and locate the Contact
  strCriteria = "[LastName] = '" & strLast & "' AND [FirstName] = '" & strFirst & "'"
  Call StartOutlook
  
  Set fdrAAContacts = GetAAContactsFolder(ns.Folders(cstPublicFoldersDir).Folders(cstAllPublicFoldersDir))
  Set itmContact = fdrAAContacts.Items.Find(strCriteria)
  
  If itmContact Is Nothing And blnCreateIfNotFound Then
    Set itmContact = fdrAAContacts.Items.Add
    itmContact.LastName = strLast
    itmContact.FirstName = strFirst
  End If

  If TypeName(itmContact) <> "Nothing" Then
    With itmContact
   
      If Len(frmUsed!txtMgrPosition & "") > 0 Then
        If .JobTitle <> frmUsed!txtMgrPosition Then
          .JobTitle = frmUsed!txtMgrPosition
          blnSave = True
        End If
      End If
      If Len(frmUsed!txtCoName & "") > 0 Then
        If .CompanyName <> frmUsed!txtCoName Then
          .CompanyName = frmUsed!txtCoName
          blnSave = True
        End If
      End If
      If Len(frmUsed!txtAddress1 & "") > 0 Then
        If .BusinessAddressStreet <> frmUsed!txtAddress1 Then
          .BusinessAddressStreet = frmUsed!txtAddress1
          blnSave = True
        End If
      End If
      If Len(frmUsed!txtAddress2 & "") > 0 Then
        If .BusinessAddressPostOfficeBox <> frmUsed!txtAddress2 Then
          .BusinessAddressPostOfficeBox = frmUsed!txtAddress2
           blnSave = True
        End If
      End If
      If Len(frmUsed!txtCity & "") > 0 Then
        If .BusinessAddressCity <> frmUsed!txtCity Then
          .BusinessAddressCity = frmUsed!txtCity
          blnSave = True
        End If
      End If
      If Len(frmUsed!txtState & "") > 0 Then
        If .BusinessAddressState <> frmUsed!txtState Then
          .BusinessAddressState = frmUsed!txtState
           blnSave = True
        End If
      End If
      If Len(frmUsed!txtZip & "") > 0 Then
        If .BusinessAddressPostalCode <> frmUsed!txtZip Then
          .BusinessAddressPostalCode = frmUsed!txtZip
           blnSave = True
        End If
      End If
      If Len(frmUsed!txtCountry & "") > 0 Then
        If .BusinessAddressCountry <> frmUsed!txtCountry Then
          .BusinessAddressCountry = frmUsed!txtCountry
          blnSave = True
        End If
      End If
      If Len(frmUsed!txtMgrEmail & "") > 0 Then
        If .Email1Address <> frmUsed!txtMgrEmail Then
          .Email1Address = frmUsed!txtMgrEmail
           blnSave = True
        End If
      End If
      If Len(frmUsed!txtMgrPhone & "") > 0 Then
        If .BusinessTelephoneNumber <> frmUsed!txtMgrPhone Then
          .BusinessTelephoneNumber = frmUsed!txtMgrPhone
          blnSave = True
        End If
      End If
      If Len(frmUsed!txtMgrFax & "") > 0 Then
        If .BusinessFaxNumber <> frmUsed!txtMgrFax Then
          .BusinessFaxNumber = frmUsed!txtMgrFax
           blnSave = True
        End If
      End If
      If blnSave = True Then .Save
    End With
  End If
  
UpdateOLContactExit:
  Call EndOutlook
  Exit Function
  
UpdateOLContactErr:
  Set itmContact = Nothing
  MsgBox Err & ": " & Err.Description, , cstMDL & " UpdateOLContact"
  Resume UpdateOLContactExit
End Function

Public Function GetOLData(frmUsed As Form, strLast As String, strFirst As String, _
                          Optional strGetWhat As String = "All")
' Possible 'strGetWhat' values (beside default 'All'): "Address", "Email/Phones"
  Dim fdrAAContacts As Outlook.MAPIFolder
  Dim itmContact As Outlook.ContactItem
  Dim strCriteria As String
  Dim blnAll As Boolean, blnEmail As Boolean, blnAdrs As Boolean
  Dim lngErr As Long
  Const cstProc = "GetOLData"
  
  strCriteria = "": blnAll = False: blnEmail = False: blnAdrs = False
  On Error GoTo GetOLDataErr
  
  If frmUsed.Name <> cstMgmtForm Then
    MsgBox "Function designed to work with the form 'frmManagement' only.", _
           vbExclamation, cstMDL & cstProc & ": wrong use"
    Exit Function
  End If
  If (Len(strLast & "") = 0 Or Len(strFirst & "") = 0) Then
    MsgBox "Both First & Last names are required.", vbExclamation, _
           cstMDL & cstProc & ": data check"
    Exit Function
  End If
  
  DoCmd.Hourglass True
  'Set flags:
  If IsMissing(strGetWhat) Then
    blnAll = True
  Else
    Select Case strGetWhat
      Case "All"
        blnAll = True
      Case "Email/Phones"
        blnEmail = True
      Case "Address"
        blnAdrs = True
      Case Else
        MsgBox "The value of the last & optional argument of GetOLData function was not recognized." & _
                 vbCrLf & "valid options: 'All' (default); 'Address'; 'Email/Phones'", vbInformation, _
                 cstMDL & cstProc
        Exit Function
    End Select
  End If

  'Establish the criteria, start OL and locate the Contact"
  strCriteria = "[LastName] = """ & strLast & """ and [FirstName] = """ & strFirst & """"
  Call StartOutlook
  Set fdrAAContacts = GetAAContactsFolder(ns.Folders(cstPublicFoldersDir).Folders(cstAllPublicFoldersDir))
  Set itmContact = fdrAAContacts.Items.Find(strCriteria)
  
  If itmContact Is Nothing Then
    MsgBox "No contact with the same first and last names was found in Outlook", _
            vbInformation, cstMDL & cstProc
    GoTo GetOLDataExit
  End If

  With itmContact
    If blnAll Or blnAdrs Then
      If Len(.JobTitle) > 0 Then
        If Not IsNull(frmUsed!txtMgrPosition) Then
          If frmUsed!txtMgrPosition <> .JobTitle Then frmUsed!txtMgrPosition = .JobTitle
        Else
          frmUsed!txtMgrPosition = .JobTitle
        End If
      End If
      
      If Len(.BusinessAddressStreet) > 0 Then
        If Not IsNull(frmUsed!txtAddress1) Then
          If frmUsed!txtAddress1 <> .BusinessAddressStreet Then frmUsed!txtAddress1 = .BusinessAddressStreet
        Else
          frmUsed!txtAddress1 = .BusinessAddressStreet
        End If
      End If

      If Len(.BusinessAddressPostOfficeBox) > 0 Then
        If Not IsNull(frmUsed!txtAddress2) Then
          If frmUsed!txtAddress2 <> .BusinessAddressPostOfficeBox Then frmUsed!txtAddress2 = .BusinessAddressPostOfficeBox
        Else
          frmUsed!txtAddress2 = .BusinessAddressPostOfficeBox
        End If
      End If
      
      If Len(.BusinessAddressCity) > 0 Then
        If Not IsNull(frmUsed!txtCity) Then
          If frmUsed!txtCity <> .BusinessAddressCity Then frmUsed!txtCity = .BusinessAddressCity
        Else
          frmUsed!txtCity = .BusinessAddressCity
        End If
      End If
      
      If Len(.BusinessAddressState) > 0 Then
        If Not IsNull(frmUsed!txtState) Then
          If frmUsed!txtState <> .BusinessAddressState Then frmUsed!txtState = .BusinessAddressState
        Else
          frmUsed!txtState = .BusinessAddressState
        End If
      End If
      
      If Len(.BusinessAddressPostalCode) > 0 Then
        If Not IsNull(frmUsed!txtZip) Then
          If frmUsed!txtZip <> .BusinessAddressPostalCode Then frmUsed!txtZip = .BusinessAddressPostalCode
        Else
          frmUsed!txtZip = .BusinessAddressPostalCode
        End If
      End If
    
      If Len(.BusinessAddressCountry) > 0 Then
        If Not IsNull(frmUsed!txtCountry) Then
          If frmUsed!txtCountry <> .BusinessAddressCountry Then frmUsed!txtCountry = .BusinessAddressCountry
        Else
          frmUsed!txtCountry = .BusinessAddressCountry
        End If
      End If
    End If 'blnAll or blnAdrs
    '
    If blnAll Or blnEmail Then
      If Len(.Email1Address) > 0 Then
        If Not IsNull(frmUsed!txtMgrEmail) Then
          If frmUsed!txtMgrEmail <> .Email1Address Then frmUsed!txtMgrEmail = .Email1Address
        Else
          frmUsed!txtMgrEmail = .Email1Address
        End If
      End If
      
      If Len(.BusinessTelephoneNumber) > 0 Then
        If Not IsNull(frmUsed!txtMgrPhone) Then
          If frmUsed!txtMgrPhone <> .BusinessTelephoneNumber Then frmUsed!txtMgrPhone = .BusinessTelephoneNumber
        Else
          frmUsed!txtMgrPhone = .BusinessTelephoneNumber
        End If
      End If
      
      If Len(.BusinessFaxNumber) > 0 Then
        If Not IsNull(frmUsed!txtMgrFax <> .BusinessFaxNumber) Then
          If frmUsed!txtMgrFax <> .BusinessFaxNumber Then frmUsed!txtMgrFax = .BusinessFaxNumber
        Else
          frmUsed!txtMgrFax = .BusinessFaxNumber
        End If
      End If
    End If  'blnAll Or blnEmail
  End With

GetOLDataExit:
  DoCmd.Hourglass False
  Set itmContact = Nothing
  Set fdrAAContacts = Nothing
  Call EndOutlook
  Exit Function
  
GetOLDataErr:
  lngErr = CLng(Err.Number)
  If lngErr <> 0 Then MsgBox lngErr & ": " & Err.Description, , cstMDL & " GetOLData"
  Resume GetOLDataExit
End Function
