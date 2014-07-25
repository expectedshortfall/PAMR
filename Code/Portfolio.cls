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

Private pCashFlows As CashFlowCollection
Private pInstruments As InstrumentCollection

Private Function SetCashFlows(valueDate As Date, MP As MarketStateProvider)
    Dim inst As IInstrument
    pCashFlows.Clear
    
    For Each inst In pInstruments
        Call pCashFlows.AddCashFlows(inst.GetCashFlows(valueDate, MP))
    Next
End Function

Public Function GetNPV(valueDate As Date, ByRef MP As MarketStateProvider) As Double
    Dim cf As CashFlow
    Dim pNPV As Double
    
    SetCashFlows valueDate, MP
    
    For Each cf In pCashFlows
        pNPV = pNPV + cf.GetNPVinPLN(valueDate, MP.GetCurrentMarketState)
    Next
    
    GetNPV = pNPV
    
End Function

Public Function GetNPVByCCY(valueDate As Date, MP As MarketStateProvider, CCY As String) As Double
    Dim cf As CashFlow
    Dim pNPV As Double
    
    SetCashFlows valueDate, MP
    
    For Each cf In pCashFlows.FilterByCCY(CCY)
        pNPV = pNPV + cf.GetNPV(valueDate, MP.GetCurrentMarketState)
    Next
    
    GetNPVByCCY = pNPV

End Function

Public Sub AddInstrument(ByRef inst As IInstrument)
    pInstruments.Add inst
End Sub

Public Sub AddInstruments(ByRef newInstruments As InstrumentCollection)
    pInstruments.AddInstruments newInstruments
End Sub

''''''''''''''''''''''
' STRUCTURAL
''''''''''''''''''''''
Private Sub Class_Initialize()
    Set pCashFlows = New CashFlowCollection
    Set pInstruments = New InstrumentCollection
End Sub

Private Sub Class_Terminate()
    Set pCashFlows = Nothing
    Set pInstruments = Nothing
End Sub
