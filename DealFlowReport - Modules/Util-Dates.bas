Attribute VB_Name = "Util-Dates"
Option Compare Database
Option Explicit
'
'*****************************************************************
' DF Reports-Util-Dates Sep-24-02 10:35
'******************************************************************
'
Const cstMdl = "Util-Dates"
'------------------------------------------------------------------------------
'The last day of the current month:     DateSerial(Year(Date()), Month(Date()) + 1, 0)
'The last day of the next month:        DateSerial(Year(Date()), Month(Date()) + 2, 0)
'The first day of the previous month:   DateSerial(Year(Date()), Month(Date())-1,1)
'The last day of the previous month:    DateSerial(Year(Date()), Month(Date()),0)
'The first day of the current quarter:  DateSerial(Year(Date()), Int((Month(Date()) - 1) / 3) * 3 + 1, 1)
'The last day of the current quarter:   DateSerial(Year(Date()), Int((Month(Date()) - 1) / 3) * 3 + 4, 0)
'The first day of the current week (Sun=1): Date() - WeekDay(Date()) + 1
'The last day of the current week:          Date() - WeekDay(Date()) + 7
'The first day of the current week (using settings in Options dialog box):  Date() - WeekDay(Date(), 0) + 1
'The last day of the current week:  Date() - WeekDay(Date(), 0) + 7
'------------------------------------------------------------------------------

Public Function GetQtrEndDate(Optional dteAnyDate As Variant) As Date
  Dim Q As Integer
  Dim dte As Date

  On Error GoTo GetQtrEndDateErr
  If Not IsMissing(dteAnyDate) Then
    dte = dteAnyDate  '?GetQtrEndDate(#2/28/1999#)
  Else
    dte = Date
  End If
  Q = DatePart("q", dte)
  
  'The last day of the calculated quarter for the given date's year:
  GetQtrEndDate = DateSerial(Year(dte), (Q * 3) + 1, 0)
      
GetQtrEndDateExit:
  Exit Function
GetQtrEndDateErr:
  MsgBox "Error (" & err & "): " & err.Description, vbExclamation, "Proc: GetQtrEndDate"
  Resume GetQtrEndDateExit
End Function

Public Function GetQtrStartDate(Optional dteAnyDate As Variant) As Date
  Dim Q As Integer
  Dim dte As Date

  On Error GoTo GetQtrStartDateErr
  If Not IsMissing(dteAnyDate) Then
    dte = dteAnyDate
  Else
    dte = Date
  End If
  
  Q = DatePart("q", dte)
  Select Case Q
    Case 1
      GetQtrStartDate = DateSerial(Year(dte), 1, 1)
    Case 2
      GetQtrStartDate = DateSerial(Year(dte), 4, 1)
    Case 3
      GetQtrStartDate = DateSerial(Year(dte), 7, 1)
    Case 4
      GetQtrStartDate = DateSerial(Year(dte), 10, 1)
    Case Else
      GetQtrStartDate = Null
  End Select
  
GetQtrStartDateExit:
  Exit Function
GetQtrStartDateErr:
  MsgBox "Error (" & err & "): " & err.Description, vbExclamation, "Proc: GetQtrStartDate"
  Resume GetQtrStartDateExit
End Function

Public Function GetGivenQtrStartDate(iQuarter As Integer, iYear As Integer) As Date

  On Error GoTo GetGivenQtrStartDateErr
  
  Select Case iQuarter
    Case 1
     GetGivenQtrStartDate = DateSerial(iYear, 1, 1)
    Case 2
     GetGivenQtrStartDate = DateSerial(iYear, 4, 1)
    Case 3
     GetGivenQtrStartDate = DateSerial(iYear, 7, 1)
    Case 4
     GetGivenQtrStartDate = DateSerial(iYear, 10, 1)
    Case Else
     GetGivenQtrStartDate = Null
  End Select
      
GetGivenQtrStartDateExit:
  Exit Function
  
GetGivenQtrStartDateErr:
  MsgBox "Error (" & err & "): " & err.Description, vbExclamation, "Proc: GetGivenQtrStartDate"
  Resume GetGivenQtrStartDateExit
End Function

Public Function GetGivenQtrEndDate(iQuarter As Integer, iYear As Integer) As Date

  On Error GoTo GetGivenQtrEndDateErr
      
  ''The last day of the given quarter for the given year:
  GetGivenQtrEndDate = DateSerial(iYear, (iQuarter * 3) + 1, 0)
  
GetGivenQtrEndDateExit:
  Exit Function
  
GetGivenQtrEndDateErr:
  MsgBox "Error (" & err & "): " & err.Description, vbExclamation, "Proc: GetGivenQtrEndDate"
  Resume GetGivenQtrEndDateExit
End Function

Public Function GetPrevQtrEndDate(Optional dteAnyDate As Variant) As Date
  Dim Q, PrevQYear, PrevQEndMonth As Integer
  Dim dte As Date
  
  On Error GoTo GetPrevQtrEndDateErr
  If Not IsMissing(dteAnyDate) Then
    dte = dteAnyDate  '?GetPrevQtrEndDate(#2/28/1999#)
  Else
    dte = Date
  End If
     
  Q = DatePart("q", dte)
  
  PrevQYear = Year(dte)  'initial value: changed for Q1
  Select Case Q   'find which quarter to get end date of prev quarter
    Case 1
      PrevQYear = Year(dte) - 1
      PrevQEndMonth = 12
    Case 2
      PrevQEndMonth = 3
    Case 3
      PrevQEndMonth = 6
    Case 4
      PrevQEndMonth = 9
  End Select
  GetPrevQtrEndDate = DateSerial(PrevQYear, PrevQEndMonth + 1, 0)
  
GetPrevQtrEndDateExit:
  Exit Function
GetPrevQtrEndDateErr:
  MsgBox "Error (" & err & "): " & err.Description, vbExclamation, "Proc: GetPrevQtrEndDate"
  Resume GetPrevQtrEndDateExit
End Function

Public Function GetNextQtrStartDate(Optional dteAnyDate As Variant) As Date
  Dim Q As Integer
  Dim dte As Date
  
  On Error GoTo GetNextQtrStartDateErr
  If Not IsMissing(dteAnyDate) Then
    dte = dteAnyDate  ' ?GetNextQtrStartDate(#2/28/1999#)
  Else
    dte = Date
  End If
  Q = Format(dte, "q")
  GetNextQtrStartDate = DateSerial(Year(dte), Int((Month(dte) - 1) / 3) * 3 + 4, 1)
      
GetNextQtrStartDateExit:
  Exit Function
GetNextQtrStartDateErr:
  MsgBox "Error (" & err & "): " & err.Description, vbExclamation, "Proc: GetNextQtrStartDate"
  Resume GetNextQtrStartDateExit
  
End Function

Function DaysInMonth(dteInput As Date) As Integer
    Dim intDays As Integer

    ' Add one month, subtract dates to find difference.
    intDays = DateSerial(Year(dteInput), Month(dteInput) + 1, Day(dteInput)) - _
              DateSerial(Year(dteInput), Month(dteInput), Day(dteInput))
    DaysInMonth = intDays
    'Debug.Print intDays
End Function

Function GetMondayDate(CurrentDate)
' Returns the date of the previous Monday
   If VarType(CurrentDate) <> vbDate Then 'vbDate=7
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

Function GetMonthEnd(iMonth As Integer, iYear As Integer) As Date
' No month range check: the year is adjusted according to month value.
  GetMonthEnd = DateSerial(iYear, iMonth + 1, 0)
End Function

Function GetMonthStart(iMonth As Integer, iYear As Integer) As Date
  GetMonthStart = DateSerial(iYear, iMonth, 1)
End Function

Function TimeToSeconds(ByVal newTime As Date) As Long
  TimeToSeconds = Hour(newTime) * 3600 + Minute(newTime) * 60 + Second(newTime)
End Function


