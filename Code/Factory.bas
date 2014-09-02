Attribute VB_Name = "Factory"
Option Explicit

Public Function CreateMarketSate(inDate As Date) As MarketState
    Set CreateMarketSate = New MarketState
    CreateMarketSate.Initialize inDate:=inDate
End Function

Public Function CreateMarketStateProvider(inDate As Date) As MarketStateProvider
    Set CreateMarketStateProvider = New MarketStateProvider
    CreateMarketStateProvider.Initialize inDate
End Function

Public Function CreateCurve(Name As String, curveDate As Date) As Curve
    Set CreateCurve = New Curve
    CreateCurve.Initialize curveDate:=curveDate, Name:=Name
End Function


Public Function CreateDiscountFactor(discountCurveName As String, dcc As DayCountConvention) As DiscountFactor
    
    Set CreateDiscountFactor = New DiscountFactor
    CreateDiscountFactor.Initialize discountCurveName:=discountCurveName, dcc:=dcc

End Function

Public Function CreateRateManager(inDiscountCurveName As String, inForwardCurveName As String, _
                                  inDcc As DayCountConvention) As RateManager
    
    Set CreateRateManager = New RateManager
    CreateRateManager.Initialize inDiscountCurveName, inForwardCurveName, inDcc

End Function

Public Function CreateFRA(inTradeDate As Date, inValueDate As Date, inMaturityDate As Date, _
                          inNominal As Double, inCurrency As String, inPosition As Position, _
                          inRate As Double, inRecFixingDate As Integer, inRateManager As RateManager) As FRA
        
        Set CreateFRA = New FRA
        CreateFRA.Initialize inTradeDate, inValueDate, inMaturityDate, inNominal, inCurrency, inPosition, _
                             inRate, inRecFixingDate, inRateManager
End Function

Public Function CreateIRS_CIRS() As IRS_CIRS
    
    Set CreateIRS_CIRS = New IRS_CIRS
    CreateIRS_CIRS.Initialize

End Function
