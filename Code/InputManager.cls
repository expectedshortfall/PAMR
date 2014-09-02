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

'=============================================================
'
'   LOADING MARKET STATE
'
'=============================================================
Public Function LoadMarketStateProvider(inDate As Date) As MarketStateProvider
    Dim msf As MarketStateFactory: Set msf = New MarketStateFactory
    
    Set LoadMarketStateProvider = msf.LoadMarketStateProvider(inDate)
    
    Set msf = Nothing
        
End Function


'=============================================================
'
'   LOADING INSTRUMENTS
'
'=============================================================
Public Function LoadPortfolio() As Portfolio
    
    Dim il As InstrumentFactory: Set il = New InstrumentFactory
    
    Set LoadPortfolio = New Portfolio
    LoadPortfolio.AddInstruments il.LoadInstruments
        
    Set il = Nothing

End Function





