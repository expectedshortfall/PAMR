VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MSP As MarketStateProvider
Private P As Portfolio
Private valuationDate As Date

'=============================================================
'
'   METHODS
'
'=============================================================
Public Function GetMarketStateProvider() As MarketStateProvider
    Set GetMarketStateProvider = MSP
End Function

Public Function GetPortfolio() As Portfolio
    Set GetPortfolio = P
End Function

Public Function GetValuationDate() As Date
    GetValuationDate = valuationDate
End Function

Public Sub LoadRiskMeasures(inRiskManager As RiskManager)
    Dim r As Range: Set r = Range("inputAreaStart")
    Dim RM As IRiskMeasure
    
    While (Not r.Value = Empty)
        inRiskManager.AddRiskMeasure Application.Run("Factory.Create" + r.Value, r.Offset(0, 1).Value)
        Set r = r.Offset(1, 0)
    Wend
End Sub

'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    Call ReadData
End Sub

Private Sub ReadData()
    
    valuationDate = Range("valuationDate").Value
    
    Set MSP = LoadMarketStateProvider(valuationDate)
    Set P = LoadPortfolio

End Sub

Private Function LoadMarketStateProvider(inDate As Date) As MarketStateProvider
    Dim msf As MarketStateFactory: Set msf = New MarketStateFactory
    
    Set LoadMarketStateProvider = msf.LoadMarketStateProvider(inDate)
        
End Function

Private Function LoadPortfolio() As Portfolio
    Dim cIF As InstrumentFactory: Set cIF = New InstrumentFactory
    
    Set LoadPortfolio = New Portfolio
    LoadPortfolio.AddInstruments cIF.LoadInstruments

End Function
