VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IInstrument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================
'
'   PUBLIC IINTERFACE INTERFACE
'
'=============================================================
Public Function GetCashFlows(valueDate As Date, MP As MarketStateProvider) As CashFlowCollection
End Function

Public Property Get Name() As String
End Property
