Attribute VB_Name = "MDL Dates"
Option Compare Database
Option Explicit
'===============================================================================
' DBFrontEnd
' MDL Dates May-16-02 11:00
'
'===============================================================================
Const cstMDL = "MDL Dates"
'------------------------------------------------------------------------------
Function GetMondayDate(CurrentDate)
  If varType(CurrentDate) <> 7 Then
     GetMondayDate = Null
  Else
    Select Case Weekday(CurrentDate)
      Case 1       ' Sunday
        GetMondayDate = CurrentDate - 6
      Case 2       ' Monday
        GetMondayDate = CurrentDate
      Case 3 To 7  ' Tuesday..Saturday
        GetMondayDate = CurrentDate - Weekday(CurrentDate) + 2
     End Select
  End If
End Function

Public Function GetPrevQtrEndDate(Optional dteAnyDate As Variant) As Date
  Dim q As Integer, prevQ As Integer
  Dim intPrevQYear As Integer, intPrevQMonth As Integer
  Dim PrevQEndDate As Date, dte As Date
  
  If Not IsMissing(dteAnyDate) Then
    If Not IsDate(dteAnyDate) Then
      MsgBox "Invalid date as argument: " & vbCrLf & _
             "defaulting to today's date.", vbCritical, "GetPrevQtrEndDate"
      Exit Function
    End If
    dte = dteAnyDate  '#2/28/1999#
  Else
    dte = Date
  End If
     
  q = Format(dte, "q")
  intPrevQYear = Year(dte)  'initial value: changed for Q1
  Select Case q   'find which quarter to get end date of prev quarter
    Case 1
      intPrevQYear = Year(dte) - 1
      intPrevQMonth = 12
    Case 2
      intPrevQMonth = 3
    Case 3
      intPrevQMonth = 6
    Case 4
      intPrevQMonth = 9
  End Select
  PrevQEndDate = DateSerial(intPrevQYear, intPrevQMonth + 1, 0)
  GetPrevQtrEndDate = PrevQEndDate
End Function

