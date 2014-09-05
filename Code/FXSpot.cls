VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FXSpot"
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
Dim pPosition As Position

Dim pNominal1 As Double
Dim pNominal2 As Double
Dim pCurrency1 As CCY
Dim pCurrency2 As CCY
Dim pRM1 As RateManager
Dim pRM2 As RateManager

'=============================================================
'
'   IMPLEMENTING IINSTRUMENT INTERFACE
'
'=============================================================
Public Function IInstrument_GetCashFlows(ValueDate As Date, MP As MarketStateProvider) As CashFlowCollection
    
    Dim res As CashFlowCollection: Set res = New CashFlowCollection
    
    If ValueDate <= pTradeDate Then
        Set IInstrument_GetCashFlows = res: Exit Function
    End If
    
    Dim cfAmount1 As Double: cfAmount1 = pNominal1 * IIf(pPosition = Buy, 1, -1)
    Dim cfAmount2 As Double: cfAmount2 = pNominal2 * IIf(pPosition = Buy, -1, 1)
    
    res.AddCashFlow Factory.CreateCashFlow(pOrigin, pCurrency1, cfAmount1, pSettlementDate, pRM1.CreateDiscountFactor)
    res.AddCashFlow Factory.CreateCashFlow(pOrigin, pCurrency2, cfAmount2, pSettlementDate, pRM2.CreateDiscountFactor)
    
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
Public Sub Initialize(inTradeDate As Date, inSettlementDate As Date, inPosition As Position, _
                      inNominal1 As Double, inCurrency1 As CCY, inRateManager1 As RateManager, _
                      inNominal2 As Double, inCurrency2 As CCY, inRateManager2 As RateManager)
                     
    pTradeDate = inTradeDate
    pSettlementDate = inSettlementDate
    pPosition = inPosition
    
    pNominal1 = inNominal1
    pNominal2 = inNominal2
    pCurrency1 = inCurrency1
    pCurrency2 = inCurrency2
    
    Set pRM1 = inRateManager1
    Set pRM2 = inRateManager2
    
End Sub

'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    pOrigin = Origin.FXSpot
    pName = "FXSpot"
End Sub

Private Sub Class_Terminate()
    Set pRM1 = Nothing
    Set pRM2 = Nothing
End Sub
