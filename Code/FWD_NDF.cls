VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FWD_NDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IInstrument

'=============================================================
'
'   PRIVATE VARIABLES
'
'=============================================================
Private pOrigin As Origin
Private pName As String

Dim pTradeDate As Date
Dim pSettlementDate As Date
Dim pFixingDate As Date
Dim pPosition As Position
Dim pNominal As Double
Dim pCCY1 As CCY
Dim pCCY2 As CCY
Dim pRate As Double
Dim pRecFixingDate As Integer
Dim pDelivery As Delivery

Dim pRM1 As RateManager
Dim pRM2 As RateManager

'=============================================================
'
'   IMPLEMENTING IINSTRUMENT INTERFACE
'
'=============================================================
Public Function IInstrument_GetCashFlows(ValueDate As Date, MSP As MarketStateProvider) As CashFlowCollection
    
    Dim res As CashFlowCollection: Set res = New CashFlowCollection
    
    If ValueDate <= pTradeDate Or ValueDate >= pSettlementDate Then
        Set IInstrument_GetCashFlows = res: Exit Function
    End If
    
    Dim forwadFXRate As Double: forwadFXRate = CalculateForwardFXRate(ValueDate, MSP)
    
    Select Case pDelivery
        Case Delivery.NDF
            res.Add Factory.CreateCashFlow(FWD_NDF, pCCY2, pPosition * pNominal * (forwadFXRate - pRate), _
                                           pSettlementDate, pRM2.CreateDiscountFactor)
        Case Delivery.Outright
            res.Add Factory.CreateCashFlow(FWD_NDF, pCCY1, pPosition * pNominal * forwadFXRate, _
                               pSettlementDate, pRM1.CreateDiscountFactor)
            
            res.Add Factory.CreateCashFlow(FWD_NDF, pCCY2, pPosition * pNominal * pRate, _
                               pSettlementDate, pRM2.CreateDiscountFactor)
        Case Else
            Err.Raise vbObject + 513, "FWD_NDF::GetCashFlows", "Unknown Delivery"
    End Select
    
    Set IInstrument_GetCashFlows = res

End Function

Public Property Get IInstrument_Name() As String
    IInstrument_Name = pName
End Property
'=============================================================
'
'   METHODS
'
'=============================================================

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Initialization, should be used only in factory module
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialize(inTradeDate As Date, inSettlementDate As Date, inPosition As Position, inNominal As Double, _
                      inCCY1 As CCY, inCCY2 As CCY, inRate As Double, inRecFixingDate As Integer, _
                      inDelivery As Delivery, inRM1 As RateManager, inRM2 As RateManager)
                     
    pTradeDate = inTradeDate
    pSettlementDate = inSettlementDate
    pFixingDate = GetFixingDate(inTradeDate, inRecFixingDate)
    pPosition = inPosition
    pNominal = inNominal
    pCCY1 = inCCY1
    pCCY2 = inCCY2
    pRate = inRate
    pRecFixingDate = inRecFixingDate
    pDelivery = inDelivery
    Set pRM1 = inRM1
    Set pRM2 = inRM2
    
End Sub

'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    pOrigin = Origin.FWD_NDF
    pName = "FWD_NDF"
End Sub

Private Sub Class_Terminate()
    Set pRM1 = Nothing
    Set pRM2 = Nothing
End Sub

Private Function CalculateForwardFXRate(ValueDate As Date, MSP As MarketStateProvider) As Double
    
    Dim a As Double
    Dim b As Double
    Dim MS As MarketState
    
    If ValueDate < pFixingDate Then
        Set MS = MSP.GetCurrentMarketState
    Else
        Set MS = MSP.GetMarketStateFromDate(pFixingDate)
    End If
    
    a = pRM2.CreateCompoundFactor.Calculate(ValueDate, pSettlementDate, MS)
    b = pRM1.CreateDiscountFactor.Calculate(ValueDate, pSettlementDate, MS)
    
    CalculateForwardFXRate = a * b * MS.GetCcyRate(pCCY1)
        
End Function


