VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Portfolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================
'
'   PRIVATE VARIABLES
'
'=============================================================
Private pCashFlows As CashFlowCollection
Private pInstruments As InstrumentCollection

'=============================================================
'
'   METHODS
'
'=============================================================
Public Function GetNPV(valueDate As Date, ByRef MP As MarketStateProvider) As Double
    
    SetCashFlows valueDate, MP
    GetNPV = DiscountCashFlows(valueDate, MP, pCashFlows)
    
End Function

Public Function GetNPVByCCY(valueDate As Date, MP As MarketStateProvider, cur As CCY) As Double
    
    SetCashFlows valueDate, MP
    GetNPVByCCY = DiscountCashFlows(valueDate, MP, pCashFlows.FilterByCCY(cur))

End Function

Public Function GetNPVByOrigin(valueDate As Date, MP As MarketStateProvider, inOrigin As Origin) As Double
        
    SetCashFlows valueDate, MP
    GetNPVByOrigin = DiscountCashFlows(valueDate, MP, pCashFlows.FilterByOrigin(inOrigin))

End Function

Public Sub AddInstrument(ByRef inst As IInstrument)
    pInstruments.Add inst
End Sub

Public Sub AddInstruments(ByRef newInstruments As InstrumentCollection)
    pInstruments.AddInstruments newInstruments
End Sub

'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Function SetCashFlows(valueDate As Date, MP As MarketStateProvider)
    Dim inst As IInstrument
    pCashFlows.Clear
    
    For Each inst In pInstruments
        Call pCashFlows.AddCashFlows(inst.GetCashFlows(valueDate, MP))
    Next
End Function

Private Function DiscountCashFlows(valueDate As Date, MP As MarketStateProvider, cashFlows As CashFlowCollection)
    Dim cf As CashFlow
    
    For Each cf In cashFlows
        DiscountCashFlows = DiscountCashFlows + cf.GetNPV(valueDate, MP.GetCurrentMarketState)
    Next

End Function

Private Sub Class_Initialize()
    Set pCashFlows = New CashFlowCollection
    Set pInstruments = New InstrumentCollection
End Sub

Private Sub Class_Terminate()
    Set pCashFlows = Nothing
    Set pInstruments = Nothing
End Sub
