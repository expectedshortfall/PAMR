VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FRA"
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
Dim pTradeDate As Date
Dim pFixingDate As Date
Dim pValueDate As Date
Dim pMaturityDate As Date
Dim pNominal As Double
Dim pCurrency As CCY
Dim pPosition As Position
Dim pRate As Double
Dim pRecFixingDate As Integer

Dim pOrigin As Origin
Dim pName As String
Dim pRateManager As RateManager

'=============================================================
'
'   IMPLEMENTING IINSTRUMENT INTERFACE
'
'=============================================================
Public Function IInstrument_GetCashFlows(ValueDate As Date, ByRef MP As MarketStateProvider) As CashFlowCollection
    
    Dim res As CashFlowCollection: Set res = New CashFlowCollection
    
    If ValueDate >= pValueDate Or ValueDate <= pTradeDate Then
        Set IInstrument_GetCashFlows = res: Exit Function
    End If
        
    Dim forwardRate As Double
    Dim cfAmount  As Double
    
    If ValueDate < pFixingDate Then
        forwardRate = pRateManager.GetForwardRate(ValueDate, pValueDate, pMaturityDate, MP.GetCurrentMarketState)
    Else
        forwardRate = pRateManager.GetRateForTenor(p3M, MP.GetMarketStateFromDate(pFixingDate))
    End If
    
    'We always buy floating rate: i.e. we recieve floating rate, and pay fixed rate
    
    Dim acc As Double: acc = Exp(0.0001 * (forwardRate - pRate) * pRateManager.GetYearFraction(pValueDate, pMaturityDate)) - 1
    Dim fdf As Double: fdf = pRateManager.GetFDF(pValueDate, pMaturityDate, forwardRate)
    
    cfAmount = acc * pNominal * fdf * IIf(pPosition = Buy, 1, -1)
    
    res.AddCashFlow Factory.CreateCashFlow(pOrigin, pCurrency, cfAmount, pValueDate, pRateManager.CreateDiscountFactor)
    
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
Public Sub Initialize(inTradeDate As Date, inValueDate As Date, inMaturityDate As Date, _
                      inNominal As Double, inCurrency As CCY, inPosition As Position, _
                      inRate As Double, inRecFixingDate As Integer, inRateManager As RateManager)
                     
    pFixingDate = GetFixingDate(inValueDate, inRecFixingDate)
    pTradeDate = inTradeDate
    pValueDate = inValueDate
    pMaturityDate = inMaturityDate
    pNominal = inNominal
    pCurrency = inCurrency
    pPosition = inPosition
    pRate = inRate
    pRecFixingDate = inRecFixingDate
    
    Set pRateManager = inRateManager
    
End Sub

'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    pOrigin = Origin.FRA
    pName = "FRA"
End Sub

Private Sub Class_Terminate()
    Set pRateManager = Nothing
End Sub
