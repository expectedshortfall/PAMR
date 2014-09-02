VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRS_CIRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IInstrument

Dim pName As String
Dim pRateManager As RateManager

'=============================================================
'
'   IMPLEMENTING IINSTRUMENT INTERFACE
'
'=============================================================
Public Function IInstrument_GetCashFlows(valueDate As Date, ByRef MP As MarketStateProvider) As CashFlowCollection
    Dim res As CashFlowCollection: Set res = New CashFlowCollection
    
    
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
Public Sub Initialize()
End Sub


'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    pName = "IRS_CIRS"
End Sub
