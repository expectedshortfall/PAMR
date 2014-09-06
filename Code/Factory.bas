Attribute VB_Name = "Factory"
Option Explicit

'=============================================================
'
'   RISK MEASURES FACTORY
'
'=============================================================
' TODO: CREATE IT PROGRAMMATICALLY

Public Function CreateBPV(inCcy As Variant) As mBPV
        
    Set CreateBPV = New mBPV
    CreateBPV.IRiskMeasure_Initialize inCcy

End Function

Public Function CreateRotation(inCcy As Variant) As mRotation
    
    Set CreateRotation = New mRotation
    CreateRotation.IRiskMeasure_Initialize inCcy

End Function

Public Function CreateStressTest(param As Variant) As mStressTest
    Set CreateStressTest = New mStressTest
End Function

Public Function CreateVaR(alpha As Double) As mVAR
    
    Set CreateVaR = New mVAR
    CreateVaR.IRiskMeasure_Initialize alpha

End Function

Public Function CreateES(alpha As Double) As mES
    
    Set CreateES = New mES
    CreateES.IRiskMeasure_Initialize alpha

End Function


'=============================================================
'
'   INSTRUMENT FACTORY
'
'=============================================================
' TODO: CREATE IT PROGRAMMATICALLY
Public Function CreateFRA(inTradeDate As Date, inValueDate As Date, inMaturityDate As Date, inNominal As Double, _
                          inCurrency As CCY, inPosition As Position, _
                          inRate As Double, inRecFixingDate As Integer, inRateManager As RateManager) As FRA
        
    Set CreateFRA = New FRA
    CreateFRA.Initialize inTradeDate, inValueDate, inMaturityDate, inNominal, inCurrency, inPosition, _
                         inRate, inRecFixingDate, inRateManager
End Function

Public Function CreateFXSpot(inTradeDate As Date, inSettlementDate As Date, inPosition As Position, _
                            inNominal1 As Double, inCurrency1 As CCY, inRateManager1 As RateManager, _
                            inNominal2 As Double, inCurrency2 As CCY, inRateManager2 As RateManager) As FXSpot
                            
    Set CreateFXSpot = New FXSpot
        CreateFXSpot.Initialize inTradeDate, inSettlementDate, inPosition, _
                                inNominal1, inCurrency1, inRateManager1, _
                                inNominal2, inCurrency2, inRateManager2
End Function

Public Function CreateFXOption(inTradeDate As Date, inMaturityDate As Date, inSettlementDate As Date, inTheta As Integer, _
                               inNominal As Double, inUnderlying As CCY, inK As Double, inPremium As Double, _
                               rm1 As RateManager, rm2 As RateManager) As FXOption

    Set CreateFXOption = New FXOption
        CreateFXOption.Initialize inTradeDate, inMaturityDate, inSettlementDate, _
                                  inTheta, inNominal, inUnderlying, inK, inPremium, _
                                  rm1, rm2

End Function

Public Function CreateFWD_NDF(inTradeDate As Date, inSettlementDate As Date, inPosition As Position, inNominal As Double, _
                              inCCY1 As CCY, inCCY2 As CCY, inRate As Double, inRecFixingDate As Integer, _
                              inDelivery As Delivery, inRM1 As RateManager, inRM2 As RateManager) As FWD_NDF

    Set CreateFWD_NDF = New FWD_NDF
        CreateFWD_NDF.Initialize inTradeDate, inSettlementDate, inPosition, inNominal, _
                      inCCY1, inCCY2, inRate, inRecFixingDate, inDelivery, _
                      inRM1, inRM2
End Function

Public Function CreateIRS_CIRS() As IRS_CIRS
    
    Set CreateIRS_CIRS = New IRS_CIRS
    CreateIRS_CIRS.Initialize

End Function

'=============================================================
'
'   MISC FACTORY
'
'=============================================================
Public Function CreateDiscountFactor(discountCurveName As String, dcc As DayCountConvention) As DiscountFactor
    
    Set CreateDiscountFactor = New DiscountFactor
    CreateDiscountFactor.Initialize discountCurveName:=discountCurveName, dcc:=dcc

End Function

Public Function CreateRateManager(inDiscountCurveName As String, inForwardCurveName As String, _
                                  inDcc As DayCountConvention) As RateManager
    
    Set CreateRateManager = New RateManager
    CreateRateManager.Initialize inDiscountCurveName, inForwardCurveName, inDcc

End Function

Public Function CreateCashFlow(inOrigin As Origin, inCcy As CCY, inValue As Double, inValueDate As Date, _
                              inDF As DiscountFactor) As CashFlow
    
    Set CreateCashFlow = New CashFlow
        CreateCashFlow.Initialize inOrigin, inCcy, inValue, inValueDate, inDF

End Function

Public Function CreateMarketSate(inDate As Date) As MarketState
    Set CreateMarketSate = New MarketState
    CreateMarketSate.Initialize inDate:=inDate
End Function

Public Function CreateMarketStateProvider(inDate As Date, Optional inMarketState As MarketState, _
                                          Optional inHistory As Scripting.Dictionary) As MarketStateProvider
    
    Set CreateMarketStateProvider = New MarketStateProvider
    CreateMarketStateProvider.Initialize inDate, inMarketState, inHistory

End Function

Public Function CreateCurve(Name As String, curveDate As Date) As Curve
    Set CreateCurve = New Curve
    CreateCurve.Initialize curveDate:=curveDate, Name:=Name
End Function



