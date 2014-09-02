Attribute VB_Name = "Types"
'=============================================================
'
'   TYPES
'
'=============================================================
Public Enum DayMoveConvention
    Following
    ModifiedFollowing
End Enum

Public Enum DayCountConvention
    cAct365
    cAct360
    c30360
    cUnknown
End Enum

Public Enum Period
    pSN
    p1M
    p2M
    p3M
    p6M
    p9M
    p1Y
    p2Y
    p3Y
    p4Y
    p5Y
    p7Y
    p10Y
    p20Y
End Enum

Public Enum CCY
    PLN
    EUR
    USD
    ALL
End Enum

Public Enum Position
    Buy
    Sell
End Enum

Public Enum Origin
    FRA
    IRS_CIRS
    FXSpot
    FXSwap
    FXOption
    FWD_NDF
    Bond
    FBond
End Enum

'=============================================================
'
'   TYPE CONVERSION
'
'=============================================================
Public Function StringToCCY(inString As String) As CCY
    If inString Like "PLN" Then
        StringToCCY = CCY.PLN
    ElseIf inString Like "EUR" Then
        StringToCCY = CCY.EUR
    ElseIf inString Like "USD" Then
        StringToCCY = CCY.USD
    Else
        StringToCCY = CCY.ALL
    End If
End Function

Public Function StringToPosition(inString As String) As Position
    If inString Like "Buy" Then
        StringToPosition = Buy
    Else
        StringToPosition = Sell
    End If
End Function

Public Function StringToPeriod(inString As String) As Period
    
    If inString Like "*SN" Or inString Like "*ON" Then
       StringToPeriod = pSN
    ElseIf inString Like "*1M" Then
        StringToPeriod = p1M
    ElseIf inString Like "*2M" Then
        StringToPeriod = p2M
    ElseIf inString Like "*3M" Then
        StringToPeriod = p3M
    ElseIf inString Like "*6M" Then
        StringToPeriod = p6M
    ElseIf inString Like "*9M" Then
        StringToPeriod = p9M
    ElseIf inString Like "*1Y" Then
        StringToPeriod = p1Y
    ElseIf inString Like "*2Y" Then
        StringToPeriod = p2Y
    ElseIf inString Like "*3Y" Then
        StringToPeriod = p3Y
    ElseIf inString Like "*4Y" Then
        StringToPeriod = p4Y
    ElseIf inString Like "*5Y" Then
        StringToPeriod = p5Y
    ElseIf inString Like "*7Y" Then
        StringToPeriod = p7Y
    ElseIf inString Like "*10Y" Then
        StringToPeriod = p10Y
    Else
        StringToPeriod = p20Y
    End If

End Function
