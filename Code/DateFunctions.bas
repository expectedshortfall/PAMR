Attribute VB_Name = "DateFunctions"
Option Explicit
'=============================================================
'
'   BUSINESS DAYS
'
'=============================================================

Public Function IsItBusinessDay(inDate As Date) As Boolean
    IsItBusinessDay = Not (IsItHoliday(inDate) Or IsItWeekend(inDate))
End Function

'=============================================================
'
'   HOLIDAYS
'
'=============================================================
Public Function IsItHoliday(inDate As Date) As Boolean
    IsItHoliday = IsItFixedHoliday(inDate) Or IsItMovingHoliday(inDate)
End Function

'=============================================================
'
'   FIXED HOLIDAYS
'
'=============================================================

Public Function IsItFixedHoliday(inDate As Date) As Boolean
        
    'New years eve
    If Month(inDate) = 1 And Day(inDate) = 1 Then
        IsItFixedHoliday = True: Exit Function
    End If
    
    'Three Kings
    If Year(inDate) >= 2011 And Month(inDate) = 6 And Day(inDate) = 6 Then
        IsItFixedHoliday = True: Exit Function
    End If
    
    'Labour Day
    If Month(inDate) = 5 And Day(inDate) = 1 Then
        IsItFixedHoliday = True: Exit Function
    End If
    
    'Constitution Day
    If Month(inDate) = 5 And Day(inDate) = 3 Then
        IsItFixedHoliday = True: Exit Function
    End If
    
    'Assumption of Mary
    If Month(inDate) = 8 And Day(inDate) = 15 Then
        IsItFixedHoliday = True: Exit Function
    End If
    
    'All Saints
    If Month(inDate) = 11 And Day(inDate) = 1 Then
        IsItFixedHoliday = True: Exit Function
    End If
    
    'Independence Day
    If Month(inDate) = 11 And Day(inDate) = 11 Then
        IsItFixedHoliday = True: Exit Function
    End If
    
    'Christmas
    If Month(inDate) = 12 And (Day(inDate) = 25 Or Day(inDate) = 26) Then
        IsItFixedHoliday = True: Exit Function
    End If
    
    IsItFixedHoliday = False
    
End Function

'=============================================================
'
'   MOVING HOLIDAYS
'
'=============================================================

Public Function IsItMovingHoliday(inDate As Date) As Boolean
    IsItMovingHoliday = (IsItEasterMonday(inDate) Or IsItCorpusChristi(inDate))
End Function

Public Function IsItEasterMonday(inDate As Date) As Boolean
    IsItEasterMonday = (inDate = (EasterUSNO(Year(inDate)) + 1))
End Function

Public Function IsItCorpusChristi(inDate As Date) As Boolean
    IsItCorpusChristi = (inDate = (EasterUSNO(Year(inDate)) + 60))
End Function

Private Function EasterUSNO(yyyy As Long) As Long
    Dim c As Long
    Dim N As Long
    Dim K As Long
    Dim i As Long
    Dim j As Long
    Dim L As Long
    Dim M As Long
    Dim D As Long
    
    c = yyyy \ 100
    N = yyyy - 19 * (yyyy \ 19)
    K = (c - 17) \ 25
    i = c - c \ 4 - (c - K) \ 3 + 19 * N + 15
    i = i - 30 * (i \ 30)
    i = i - (i \ 28) * (1 - (i \ 28) * (29 \ (i + 1)) * ((21 - N) \ 11))
    j = yyyy + yyyy \ 4 + i + 2 - c + c \ 4
    j = j - 7 * (j \ 7)
    L = i - j
    M = 3 + (L + 40) \ 44
    D = L + 28 - 31 * (M \ 4)
    EasterUSNO = DateSerial(yyyy, M, D)
End Function

'=============================================================
'
'   WEEKEND
'
'=============================================================
Public Function IsItWeekend(dtmTemp As Date) As Boolean
    Select Case Weekday(dtmTemp)
        Case vbSaturday, vbSunday
            IsItWeekend = True
        Case Else
            IsItWeekend = False
    End Select
End Function

'=============================================================
'
'   YEAR FRACTIONS
'
'=============================================================
Public Function YearFrac(startDate As Date, endDate As Date, convention As DayCountConvention) As Double
    Select Case convention
        Case DayCountConvention.cAct365
            YearFrac = dccACT365(startDate, endDate)
        Case DayCountConvention.cAct360
            YearFrac = dccACT360(startDate, endDate)
        Case DayCountConvention.c30360
            YearFrac = dcc30360(startDate, endDate)
        Case DayCountConvention.cUnknown
            YearFrac = -1
    End Select
End Function

'
' ACT/365 day count convention
'
Public Function YearFracStr(startDate As Date, endDate As Date, convention As String) As Double
    Dim tmpString As String: tmpString = Trim(convention)
    
    If InStr(tmpString, "/") Then
        convention = Left(tmpString, InStr(tmpString, "/") - 1) + Right(tmpString, Len(tmpString) - InStr(tmpString, "/"))
    End If
            
    YearFracStr = Application.Run("dcc" + convention, startDate, endDate)
    
End Function

'
' ACT/365 day count convention
'
Public Function dccACT365(startDate As Date, endDate As Date) As Double
    dccACT365 = (endDate - startDate) / 365
End Function

'
' ACT/360 day count convention
'
Public Function dccACT360(startDate As Date, endDate As Date) As Double
    dccACT360 = (endDate - startDate) / 360
End Function

'
' 30/360 day count convention
'
Public Function dcc30360(startDate As Date, endDate As Date) As Double
            
    Dim DDi As Integer: DDi = Application.WorksheetFunction.Min(Day(startDate), 30)
    Dim JJi As Integer

    If DDi = 30 Then
        JJi = Application.WorksheetFunction.Min(Day(endDate), 30)
    Else
        JJi = Day(endDate)
    End If
        
    dcc30360 = (360 * (Year(endDate) - Year(startDate)) + 30 * (Month(endDate) - Month(startDate)) + JJi - DDi) / 360

End Function

'=============================================================
'
'   DATE SHIFTS
'
'=============================================================
Public Function GetFixingDate(ValueDate As Date, shift As Integer)
    Dim i As Integer
    Dim res As Date: res = ValueDate
    
    For i = 1 To -shift
        res = PreviousBusinessDay(res)
    Next i
    
    GetFixingDate = res

End Function

Public Function CurveTenorDate(inDate As Date, inPeriod As Period)
    
    Dim tmpDate As Date: tmpDate = DateMove(inDate, pSN, Following)
    
    If Not inPeriod = pSN Then
        tmpDate = DateMove(tmpDate, inPeriod, Following)
    End If
    
    CurveTenorDate = tmpDate
    
End Function

Public Function DateMoveStr(inDate As Date, inPeriod As String, periodCount As Integer, convention As DayMoveConvention)
    DateMoveStr = ShiftIfItIsHoliday(DateAdd(inPeriod, periodCount, inDate), convention)
End Function

Public Function DateMove(inDate As Date, inPeriod As Period, convention As DayMoveConvention)
       
    Select Case inPeriod
        Case pSN
            DateMove = NexBusinessDay(NexBusinessDay(inDate))
        Case p1M
            DateMove = ShiftIfItIsHoliday(DateAdd("m", 1, inDate), convention)
        Case p2M
            DateMove = ShiftIfItIsHoliday(DateAdd("m", 2, inDate), convention)
        Case p3M
            DateMove = ShiftIfItIsHoliday(DateAdd("m", 3, inDate), convention)
        Case p6M
            DateMove = ShiftIfItIsHoliday(DateAdd("m", 6, inDate), convention)
        Case p9M
            DateMove = ShiftIfItIsHoliday(DateAdd("m", 9, inDate), convention)
        Case p1Y
            DateMove = ShiftIfItIsHoliday(DateAdd("yyyy", 1, inDate), convention)
        Case p2Y
            DateMove = ShiftIfItIsHoliday(DateAdd("yyyy", 2, inDate), convention)
        Case p3Y
            DateMove = ShiftIfItIsHoliday(DateAdd("yyyy", 3, inDate), convention)
        Case p4Y
            DateMove = ShiftIfItIsHoliday(DateAdd("yyyy", 4, inDate), convention)
        Case p5Y
            DateMove = ShiftIfItIsHoliday(DateAdd("yyyy", 5, inDate), convention)
        Case p7Y
            DateMove = ShiftIfItIsHoliday(DateAdd("yyyy", 7, inDate), convention)
        Case p10Y
            DateMove = ShiftIfItIsHoliday(DateAdd("yyyy", 10, inDate), convention)
        Case p20Y
            DateMove = ShiftIfItIsHoliday(DateAdd("yyyy", 20, inDate), convention)
    End Select

End Function

Private Function ShiftIfItIsHoliday(inDate As Date, convention As DayMoveConvention)
    If Not IsItBusinessDay(inDate) Then
        ShiftIfItIsHoliday = ShiftDate(inDate, convention)
    Else
        ShiftIfItIsHoliday = inDate
    End If
        
End Function

Public Function ShiftDate(inDate As Date, convention As DayMoveConvention) As Date
    Select Case convention
        Case DayMoveConvention.Following
            ShiftDate = ShiftDateUsingFollowing(inDate)
        Case DayMoveConvention.ModifiedFollowing
            ShiftDate = ShiftDateUsingModifiedFollowing(inDate)
    End Select
End Function

Public Function ShiftDateUsingFollowing(inDate As Date) As Date
    ShiftDateUsingFollowing = NexBusinessDay(inDate)
End Function

Public Function ShiftDateUsingModifiedFollowing(inDate As Date) As Date
    Dim resDate As Date: resDate = NexBusinessDay(inDate)
            
    If Not Month(inDate) = Month(resDate) Then
        resDate = PreviousBusinessDay(inDate)
    End If
        
    ShiftDateUsingModifiedFollowing = resDate
End Function

'
' Calculates next business Day
'
Public Function NexBusinessDay(inDate As Date) As Date
    Dim resDate As Date: resDate = inDate + 1
        
    While Not IsItBusinessDay(resDate)
        resDate = resDate + 1
    Wend
    
    NexBusinessDay = resDate

End Function

'
' Calculates Previous Business Day
'
Public Function PreviousBusinessDay(inDate As Date) As Date
    Dim resDate As Date: resDate = inDate - 1
        
    While Not IsItBusinessDay(resDate)
        resDate = resDate - 1
    Wend
    
    PreviousBusinessDay = resDate

End Function



