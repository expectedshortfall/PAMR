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
    Buy = 1
    Sell = -1
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

Public Enum Delivery
    NDF
    Outright
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

Public Function CcyToString(inCcy As CCY) As String
    Select Case inCcy
        Case CCY.PLN
            CcyToString = "PLN"
        Case CCY.EUR
            CcyToString = "EUR"
        Case CCY.USD
            CcyToString = "USD"
        Case CCY.ALL
            CcyToString = "ALL"
        Case Else
        Err.Raise vbObject + 513, "Types::CcyToString", "Unknown currency"
    End Select
End Function

Public Function StringToPosition(inString As String) As Position
    If inString Like "Buy" Then
        StringToPosition = Buy
    Else
        StringToPosition = Sell
    End If
End Function

Public Function StringToDelivery(inString As String) As Delivery
    If inString Like "NDF" Then
        StringToDelivery = NDF
    Else
        StringToDelivery = Outright
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

Public Function CastToCurve(curveObj As Variant) As Curve
    Set CastToCurve = curveObj
End Function

Public Function CastToMarketState(marketStateObj As Variant) As MarketState
    Set CastToMarketState = marketStateObj
End Function
