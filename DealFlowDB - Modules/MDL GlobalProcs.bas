Attribute VB_Name = "MDL GlobalProcs"
Option Compare Database
Option Explicit
'================================================================================
' DBFrontEnd
' MDL GlobalProcs Sep-05-03 14:40
'
'================================================================================
Public dbs As DAO.Database
'--------------------------------------------------------------------------------
Public Const cstMainForm = "frmDealSelection"
Public Const cstDealForm = "frmDeal"
Public Const cstSourceForm = "frmSource"
Public Const cstIssuerForm = "frmIssuer"
Public Const cstFinStatusForm = "frmFinStat"
Public Const cstMgmtForm = "frmManagement"
Public Const cstFilterForm = "frmFilterForm"
Public Const cstReportForm = "frmReportSelection"
Public Const cstDataForm = "frmDatasheet"
Public Const cstDataSubForm = "frmDatasheetSub"
'--------------------------------------------------------------------------------
Public Const cstSaveMsg = "Do you want to save your changes?"
'--------------------------------------------------------------------------------
Public strCallingForm As String   ' Set/reset when a form is opened
Public cbxCallingBox As ComboBox
Public blnFilterErr As Boolean
Public blnNewDeal As Boolean
Public blnNewIssuer As Boolean
Public blnNewSource As Boolean
'------------------------------------------------------------------------------------------
Public Const cstAAColor = 5785120: Public Const cstAAColorLight = 10325248
'------------------------------------------------------------------------------------------

Public Function IsSet(var As Variant) As Boolean
  Dim obj As Object
  Dim str As String
  IsSet = False
  On Error GoTo IsSetErr
  If Not IsObject(var) Then Exit Function
  Set obj = var
  str = obj.Name
  IsSet = True

IsSetExit:
  Set obj = Nothing
  Exit Function
  
IsSetErr:
  IsSet = False
  Resume IsSetExit
End Function

Function IsLoaded(ByVal strformName As String) As Boolean
  Const conObjStateClosed = 0
  Const conDesignView = 0
  If SysCmd(acSysCmdGetObjectState, acForm, strformName) <> conObjStateClosed Then
    If Forms(strformName).CurrentView <> conDesignView Then IsLoaded = True
  End If
End Function

Function RptNoData() As Integer
  MsgBox "No Data To Preview", vbExclamation, "Operation Cancelled"
  RptNoData = True
End Function

Function CloseAllOpenFormsButCaller(frmCallingForm As Form)
  Dim axsObj As AccessObject
  Dim str As String
  ' Search for open AccessObject objects in AllForms collection.
  For Each axsObj In Application.CurrentProject.AllForms
    If axsObj.IsLoaded = True Then
      str = axsObj.Name
      If str <> frmCallingForm.Name Then DoCmd.Close acForm, str
    End If
  Next axsObj
  Set axsObj = Nothing
  DoCmd.Hourglass False
End Function
